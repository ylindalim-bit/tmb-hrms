# -*- coding: utf-8 -*-
"""
Payroll calculation engine.

Every formula here is a direct port of the corresponding cell formula in
Payroll_Master_2026_9.xlsm ('Payroll' sheet columns F-X, and
'PCB Inputs & YTD' sheet columns I-T for the LHDN Computerised Method PCB
calculation) - not re-derived from general LHDN/KWSP guidance. Comments
reference the source column letters so this can be checked against the
workbook cell-by-cell.
"""
import calendar
import datetime
import sqlite3

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul",
              "Aug", "Sep", "Oct", "Nov", "Dec"]


def _parse_date(s):
    if not s:
        return None
    return datetime.date.fromisoformat(s)


def allowance_prorate_factor(effective_date, year, month):
    """
    Fraction of (year, month) during which an allowance with this effective
    date is active - same day-count logic as prorate_factor() above, but
    anchored to the allowance's own effective date instead of date joined.
    No effective date set = always fully active (1.0). Effective date after
    the month = not yet active (0.0). Effective date mid-month = prorated
    by the days remaining in the month from that date.
    """
    d = _parse_date(effective_date)
    if d is None:
        return 1.0
    days_in_month = calendar.monthrange(year, month)[1]
    month_start = datetime.date(year, month, 1)
    month_end = datetime.date(year, month, days_in_month)
    if d <= month_start:
        return 1.0
    if d > month_end:
        return 0.0
    days_active = (month_end - d).days + 1
    return days_active / days_in_month


def prorate_factor(date_joined, last_working_day, year, month):
    """Payroll!L and Payroll!M: calendar-day prorate factor for the month."""
    days_in_month = calendar.monthrange(year, month)[1]
    month_start = datetime.date(year, month, 1)
    month_end = datetime.date(year, month, days_in_month)

    joined = _parse_date(date_joined) or month_start
    left = _parse_date(last_working_day) or month_end

    effective_end = min(month_end, left)
    effective_start = max(month_start, joined)

    days_employed = max(0, (effective_end - effective_start).days + 1)
    factor = days_employed / days_in_month if days_in_month else 0
    return days_employed, factor


def calc_ot_pay(basic_salary, fixed_allowance, working_hours_day,
                 ot_hours_1_5, ot_hours_2_0, ot_hours_3_0):
    """
    Hourly rate = (Basic+Fixed)/26/WorkingHoursPerDay (per the workbook's
    ORP formula), applied across the three OT rate tiers from the
    handbook's Figure 2.3(a) overtime table:
      - 1.5x: normal working days and Saturdays, up to ~10pm
      - 2.0x: normal days/Saturdays after ~10pm, and public holidays up to ~10pm
      - 3.0x: public holidays after ~10pm
    Which hours fall in which tier is a judgment call for whoever enters
    attendance (it depends on actual clock times, which this app doesn't
    capture) - the three fields on the Attendance page let them assign
    approved OT hours to the correct tier directly.
    """
    if not working_hours_day:
        return 0.0, 0.0, 0.0, 0.0
    hourly_rate = (basic_salary + fixed_allowance) / 26 / working_hours_day
    ot_pay_1_5 = round(ot_hours_1_5 * 1.5 * hourly_rate, 2)
    ot_pay_2_0 = round(ot_hours_2_0 * 2.0 * hourly_rate, 2)
    ot_pay_3_0 = round(ot_hours_3_0 * 3.0 * hourly_rate, 2)
    return ot_pay_1_5, ot_pay_2_0, ot_pay_3_0, round(hourly_rate, 4)


def _vlookup_bracket_label(rows, wage):
    """Same matching rule as _vlookup_bracket, but returns the human-readable
    bracket_label of the matched row - used to explain which official wage
    band a contribution was looked up from."""
    match = rows[0] if rows else None
    for row in rows:
        if row["wage_lower_bound"] <= wage:
            match = row
        else:
            break
    return match["bracket_label"] if match else None


def _vlookup_bracket(rows, wage, key_field):
    """Excel VLOOKUP(wage, table, col, TRUE): last row whose lower bound <= wage."""
    match = rows[0] if rows else None
    for row in rows:
        if row["wage_lower_bound"] <= wage:
            match = row
        else:
            break
    return match[key_field] if match else 0.0


def lookup_epf(wage_base, age, epf_rows):
    """Payroll!O (employee) and Payroll!P (employer)."""
    if wage_base > 20000:
        if age < 60:
            employee = -(-wage_base * 0.11 // 1)  # CEILING(...,1)
            employer = -(-wage_base * 0.12 // 1)
        else:
            employee = -(-wage_base * 0 // 1)
            employer = -(-wage_base * 0.04 // 1)
        return float(employee), float(employer)
    if age < 60:
        employee = _vlookup_bracket(epf_rows, wage_base, "part_a_under60_employee")
        employer = _vlookup_bracket(epf_rows, wage_base, "part_a_under60_employer")
    else:
        employee = _vlookup_bracket(epf_rows, wage_base, "part_e_60plus_employee")
        employer = _vlookup_bracket(epf_rows, wage_base, "part_e_60plus_employer")
    return employee, employer


def lookup_socso(gross_pay, age, socso_rows):
    """Payroll!Q (employee) and Payroll!R (employer). No wages paid this
    month (e.g. fully on Unpaid Leave) means no contribution - the "Up to
    RM30" bracket's small non-zero amount is for someone who actually earned
    something in that range, not RM0."""
    if gross_pay <= 0:
        return 0.0, 0.0
    if age < 60:
        employee = _vlookup_bracket(socso_rows, gross_pay, "cat1_employee_invalidity")
        employer = _vlookup_bracket(socso_rows, gross_pay, "cat1_employer")
    else:
        employee = 0.0
        employer = _vlookup_bracket(socso_rows, gross_pay, "cat2_employer")
    return employee, employer


def lookup_skbbk(gross_pay, age, socso_rows):
    """
    SKBBK employee contribution, by wage bracket and age (same Cat 1/Cat 2
    split as regular SOCSO). Applied unconditionally to every employee -
    confirmed against a real finalized payroll run (June 2026), where every
    employee was charged SKBBK matching this table exactly. The original
    workbook gated this behind an Employee Master flag that was never
    actually set to "Y" for any employee, silently zeroing it out; that
    flag is not used here.

    No wages paid this month means no contribution, same reasoning as
    lookup_socso above.
    """
    if gross_pay <= 0:
        return 0.0
    if age < 60:
        return _vlookup_bracket(socso_rows, gross_pay, "cat1_employee_skbbk")
    return _vlookup_bracket(socso_rows, gross_pay, "cat2_employee_skbbk")


def calc_eis(gross_pay, age, eis_rows, eis_flag="Y"):
    """
    PERKESO EIS (Act 800) contribution, looked up from the official bracket
    table (employer and employee always pay the identical amount at every
    bracket - not a flat percentage, despite the nominal "0.2%" often
    quoted). No contribution outside the 18-59 working-age range, if
    eis_flag has been manually set to "N" (e.g. an exempted employee), or
    if no wages were paid this month (e.g. fully on Unpaid Leave) - same
    reasoning as lookup_socso.
    """
    if eis_flag == "N" or age >= 60 or age < 18 or gross_pay <= 0:
        return 0.0, 0.0
    amount = _vlookup_bracket(eis_rows, gross_pay, "contribution")
    return round(amount, 2), round(amount, 2)


def calc_hrd_levy(gross_pay, registered, rate):
    """Payroll!AB - employer-only cost, not deducted from employee."""
    if registered != "Y":
        return 0.0
    return round(gross_pay * rate, 2)


def calc_pcb(month_index, current_gross, current_epf_employee, current_socso_eis_employee,
             ytd_gross, ytd_epf_employee, ytd_pcb_this_employer,
             tax_category, children_full, children_half, tp1_submitted, zakat_ytd,
             pcb_brackets, pcb_constants):
    """
    'PCB Inputs & YTD' columns I-T: LHDN Computerised Calculation Method for
    NORMAL monthly remuneration only (not one-off bonuses).
    """
    months_remaining = 13 - month_index  # J

    est_annual_remun = ytd_gross + current_gross * months_remaining  # N
    epf_relief = min(
        ytd_epf_employee + current_epf_employee * months_remaining,
        pcb_constants["epf_relief_cap"],
    )  # O
    socso_eis_relief = min(
        current_socso_eis_employee * months_remaining,
        pcb_constants["socso_eis_relief_cap"],
    )  # P

    spouse_child_relief = 0.0  # Q
    if tp1_submitted == "Y":
        if tax_category == "Married":
            spouse_child_relief += pcb_constants["spouse_relief"]
        spouse_child_relief += children_full * pcb_constants["child_relief_full"]
        spouse_child_relief += children_half * pcb_constants["child_relief_half"]

    chargeable_income = max(
        0.0,
        est_annual_remun
        - pcb_constants["individual_relief"]
        - epf_relief
        - socso_eis_relief
        - spouse_child_relief,
    )  # R

    bracket = pcb_brackets[0]
    for b in pcb_brackets:
        if b["income_from"] <= chargeable_income:
            bracket = b
        else:
            break
    tax_before_rebate = (
        (chargeable_income - bracket["income_from"]) * bracket["rate"] + bracket["base_tax"]
    )

    rebate = 0.0
    if chargeable_income <= pcb_constants["rebate_threshold"]:
        rebate += pcb_constants["rebate_amount"]
        if tax_category == "Married" and tp1_submitted == "Y":
            rebate += pcb_constants["rebate_amount"]

    annual_tax = max(0.0, tax_before_rebate - rebate)  # S

    pcb_this_month = max(
        0.0,
        round(
            (annual_tax - zakat_ytd - ytd_pcb_this_employer) / months_remaining, 2
        ),
    )  # T
    return {
        "months_remaining": months_remaining,
        "est_annual_remun": round(est_annual_remun, 2),
        "epf_relief": round(epf_relief, 2),
        "socso_eis_relief": round(socso_eis_relief, 2),
        "spouse_child_relief": round(spouse_child_relief, 2),
        "chargeable_income": round(chargeable_income, 2),
        "bracket_income_from": bracket["income_from"],
        "bracket_rate": bracket["rate"],
        "bracket_base_tax": bracket["base_tax"],
        "tax_before_rebate": round(tax_before_rebate, 2),
        "rebate": round(rebate, 2),
        "annual_tax": round(annual_tax, 2),
        "ytd_pcb_this_employer": ytd_pcb_this_employer,
        "zakat_ytd": zakat_ytd,
        "pcb_this_month": pcb_this_month,
    }


def calculate_payroll(conn: sqlite3.Connection, emp_id: str, year: int, month: int,
                       variable_allowance: float = None, variable_allowance_flag: str = None):
    conn.row_factory = sqlite3.Row
    emp = conn.execute("SELECT * FROM employees WHERE emp_id=?", (emp_id,)).fetchone()
    if emp is None:
        raise ValueError(f"Unknown employee {emp_id}")

    adj = conn.execute(
        "SELECT * FROM monthly_adjustments WHERE emp_id=? AND year=? AND month=?",
        (emp_id, year, month),
    ).fetchone()
    if variable_allowance is None:
        variable_allowance = adj["variable_allowance"] if adj else 0.0
    if variable_allowance_flag is None:
        variable_allowance_flag = adj["variable_allowance_flag"] if adj else "N"
    other_deduction = (adj["other_deduction"] if adj else 0.0) or 0.0
    other_deduction_desc = (adj["other_deduction_desc"] if adj else "") or ""

    att = conn.execute(
        "SELECT * FROM attendance_monthly WHERE emp_id=? AND year=? AND month=?",
        (emp_id, year, month),
    ).fetchone()
    working_days_in_month = att["working_days_in_month"] if att else 0
    unpaid_days = (att["ul_days"] + att["absent_days"]) if att else 0
    paid_leave_days = (
        (att["al_days"] + att["mc_days"] + att["hl_days"] + att["other_paid_leave"])
        if att else 0
    )
    ot_hours_1_5 = att["ot_hours_1_5"] if att else 0
    ot_hours_2_0 = att["ot_hours_2_0"] if att else 0
    ot_hours_3_0 = att["ot_hours_3_0"] if att else 0
    ot_hours = ot_hours_1_5 + ot_hours_2_0 + ot_hours_3_0
    meal_eligible_days = att["meal_eligible_days"] if att else 0
    cewi_eligible_days = att["cewi_eligible_days"] if att else 0

    settings = {r["key"]: r["value"] for r in conn.execute("SELECT * FROM payroll_settings")}
    hrd_registered = settings.get("hrd_levy_registered", "N")
    hrd_rate = float(settings.get("hrd_levy_rate", 0.01))

    def allowance_month_factor(flag, effective_date):
        if flag != "Y":
            return 0.0
        return allowance_prorate_factor(effective_date, year, month)

    basic_salary = emp["basic_salary"] or 0
    # Full (unprorated) Fixed Allowance - used for the OT hourly-rate (ORP)
    # calc below, matching the workbook's formula which uses the flagged
    # amount as-is, not a prorated one.
    fixed_allowance = (
        (emp["position_allowance"] if emp["position_allowance_flag"] == "Y" else 0)
        + (emp["training_incentive"] if emp["training_incentive_flag"] == "Y" else 0)
        + (emp["oversea_incentive"] if emp["oversea_incentive_flag"] == "Y" else 0)
    )
    # Same components, but each prorated by its own effective-date factor -
    # used for Gross Pay so a mid-month allowance start doesn't pay a full
    # month's worth.
    fixed_allowance_prorated = (
        (emp["position_allowance"] or 0) * allowance_month_factor(emp["position_allowance_flag"], emp["position_allowance_effective_date"])
        + (emp["training_incentive"] or 0) * allowance_month_factor(emp["training_incentive_flag"], emp["training_incentive_effective_date"])
        + (emp["oversea_incentive"] or 0) * allowance_month_factor(emp["oversea_incentive_flag"], emp["oversea_incentive_effective_date"])
    )

    days_employed, factor = prorate_factor(emp["date_joined"], emp["last_working_day"], year, month)

    ot_pay_1_5, ot_pay_2_0, ot_pay_3_0, ot_hourly_rate = calc_ot_pay(
        basic_salary, fixed_allowance, emp["working_hours_day"] or 8,
        ot_hours_1_5, ot_hours_2_0, ot_hours_3_0)
    ot_pay = round(ot_pay_1_5 + ot_pay_2_0 + ot_pay_3_0, 2)

    transport_allowance = (
        (emp["transport_allowance"] or 0) * factor
        * allowance_month_factor(emp["transport_allowance_flag"], emp["transport_allowance_effective_date"])
    )
    meal_allowance = round(
        meal_eligible_days * (emp["meal_allowance_rate"] or 0)
        * allowance_month_factor(emp["meal_allowance_flag"], emp["meal_allowance_effective_date"]), 2
    )
    cewi_allowance = round(
        cewi_eligible_days * (emp["cewi_rate"] or 0)
        * allowance_month_factor(emp["cewi_flag"], emp["cewi_effective_date"]), 2
    )

    prorated_basic = basic_salary * factor
    # Unpaid Leave / Absent deduction uses the number of calendar days in
    # the month (not "Working Days in Month", which is a separate figure
    # HR enters for attendance tracking) - e.g. August has 31 days
    # regardless of how many of them are actual working days.
    calendar_days_in_month = calendar.monthrange(year, month)[1]
    unpaid_deduction = round((prorated_basic / calendar_days_in_month) * unpaid_days, 2)
    basic_after_unpaid = prorated_basic - (prorated_basic / calendar_days_in_month) * unpaid_days

    gross_pay = (
        basic_after_unpaid
        + fixed_allowance_prorated * factor
        + (variable_allowance if variable_allowance_flag == "Y" else 0.0)
        + ot_pay
        + transport_allowance
        + meal_allowance
        + cewi_allowance
    )

    dob = _parse_date(emp["date_of_birth"])
    today = datetime.date(year, month, min(calendar.monthrange(year, month)[1], 28))
    age = (today - dob).days // 365 if dob else 30

    epf_wage_base = gross_pay - ot_pay - transport_allowance
    epf_rows = [dict(r) for r in conn.execute("SELECT * FROM epf_table ORDER BY wage_lower_bound")]
    epf_employee, epf_employer = lookup_epf(epf_wage_base, age, epf_rows)
    epf_bracket_label = _vlookup_bracket_label(epf_rows, epf_wage_base)
    # Employee-elected voluntary EPF on top of the statutory rate, stored as
    # a percentage of the same EPF wage base as the statutory contribution.
    # Employer share is unaffected; this still counts toward the same EPF
    # relief cap for PCB, so it's folded in before the PCB calculation below.
    additional_epf_pct = emp["additional_epf_employee"] or 0
    additional_epf = round(epf_wage_base * (additional_epf_pct / 100), 2)
    epf_employee += additional_epf

    socso_rows = [dict(r) for r in conn.execute("SELECT * FROM socso_table ORDER BY wage_lower_bound")]
    if gross_pay <= 0:
        socso_gate_reason = "No wages paid this month (gross pay RM0) - no contribution due"
    else:
        socso_gate_reason = None
    socso_employee, socso_employer = lookup_socso(gross_pay, age, socso_rows)
    socso_bracket_label = _vlookup_bracket_label(socso_rows, gross_pay) if gross_pay > 0 else None
    # SKBBK is a new deduction the company only started charging from
    # June 2026 onward (confirmed absent from Jan/Feb/Apr/May payroll
    # records) - gated by a start date, not applied retroactively.
    skbbk_start = settings.get("skbbk_start_date")
    skbbk_applies = emp["skbbk_flag"] != "N"
    if not skbbk_applies:
        skbbk_gate_reason = "Employee's SKBBK flag is set to N"
    elif skbbk_start and f"{year:04d}-{month:02d}-01" < skbbk_start:
        skbbk_gate_reason = f"SKBBK not yet in effect this month (starts {skbbk_start})"
    elif gross_pay <= 0:
        skbbk_gate_reason = "No wages paid this month (gross pay RM0) - no contribution due"
    else:
        skbbk_gate_reason = None
    if skbbk_gate_reason is None:
        skbbk_employee = lookup_skbbk(gross_pay, age, socso_rows)
        skbbk_bracket_label = _vlookup_bracket_label(socso_rows, gross_pay)
    else:
        skbbk_employee = 0.0
        skbbk_bracket_label = None

    eis_rows = [dict(r) for r in conn.execute("SELECT * FROM eis_table ORDER BY wage_lower_bound")]
    if emp["eis_flag"] == "N":
        eis_gate_reason = "Employee's EIS flag is set to N"
    elif age >= 60 or age < 18:
        eis_gate_reason = f"Outside the 18-59 EIS working-age range (age {age})"
    elif gross_pay <= 0:
        eis_gate_reason = "No wages paid this month (gross pay RM0) - no contribution due"
    else:
        eis_gate_reason = None
    eis_employee, eis_employer = calc_eis(gross_pay, age, eis_rows, emp["eis_flag"])
    eis_bracket_label = _vlookup_bracket_label(eis_rows, gross_pay) if eis_gate_reason is None else None

    hrd_levy = calc_hrd_levy(gross_pay, hrd_registered, hrd_rate)

    # ---- PCB ----
    tax_profile = conn.execute("SELECT * FROM tax_profile WHERE emp_id=?", (emp_id,)).fetchone()
    pcb_constants = {r["key"]: r["value"] for r in conn.execute("SELECT * FROM pcb_constants")}
    pcb_brackets = [dict(r) for r in
                     conn.execute("SELECT * FROM pcb_tax_brackets ORDER BY income_from")]

    ytd_row = conn.execute(
        """SELECT COALESCE(SUM(gross_remun),0) AS gross, COALESCE(SUM(epf_employee),0) AS epf,
                  COALESCE(SUM(pcb_deducted),0) AS pcb
           FROM pcb_monthly_record WHERE emp_id=? AND year=? AND month<?""",
        (emp_id, year, month),
    ).fetchone()
    # TP3: income/EPF/PCB the employee declared from a PREVIOUS employer
    # earlier this same year (e.g. joined mid-year) - folded into the YTD
    # figures so the Computerised Method sees their true full-year income,
    # not just what's been paid through this employer.
    tp3_gross = (tax_profile["tp3_prior_gross"] or 0) if tax_profile else 0
    tp3_epf = (tax_profile["tp3_prior_epf_employee"] or 0) if tax_profile else 0
    tp3_pcb = (tax_profile["tp3_prior_pcb"] or 0) if tax_profile else 0

    pcb_breakdown = calc_pcb(
        month_index=month,
        current_gross=gross_pay,
        current_epf_employee=epf_employee,
        current_socso_eis_employee=socso_employee + eis_employee,
        ytd_gross=ytd_row["gross"] + tp3_gross,
        ytd_epf_employee=ytd_row["epf"] + tp3_epf,
        ytd_pcb_this_employer=ytd_row["pcb"] + tp3_pcb,
        tax_category=(tax_profile["tax_category"] if tax_profile else "Single"),
        children_full=(tax_profile["children_full_relief"] if tax_profile else 0),
        children_half=(tax_profile["children_half_relief"] if tax_profile else 0),
        tp1_submitted=(tax_profile["tp1_submitted"] if tax_profile else ""),
        zakat_ytd=(tax_profile["zakat_paid_ytd"] if tax_profile else 0),
        pcb_brackets=pcb_brackets,
        pcb_constants=pcb_constants,
    )
    pcb = pcb_breakdown["pcb_this_month"]

    # A manual PCB correction (e.g. from the employee's own LHDN e-PCB slip)
    # sticks across re-Finalizing: once set on a prior run for this
    # emp/year/month, it keeps overriding the calculated pcb here too.
    override_row = conn.execute(
        "SELECT pcb_override, pcb_override_reason FROM payroll_runs WHERE emp_id=? AND year=? AND month=?",
        (emp_id, year, month),
    ).fetchone()
    pcb_override = override_row["pcb_override"] if override_row else None
    pcb_override_reason = override_row["pcb_override_reason"] if override_row else None
    if pcb_override is not None:
        pcb = pcb_override

    total_deductions = epf_employee + socso_employee + eis_employee + pcb + skbbk_employee + other_deduction
    net_pay = gross_pay - total_deductions

    return {
        "emp_id": emp_id,
        "full_name": emp["full_name"],
        "year": year,
        "month": month,
        "basic_salary": round(basic_after_unpaid, 2),
        "fixed_allowance": round(fixed_allowance_prorated * factor, 2),
        "variable_allowance": round(variable_allowance, 2),
        "working_days_in_month": working_days_in_month,
        "paid_leave_days": paid_leave_days,
        "unpaid_days": unpaid_days,
        "unpaid_deduction": unpaid_deduction,
        "ot_hours": ot_hours,
        "ot_hours_1_5": ot_hours_1_5,
        "ot_hours_2_0": ot_hours_2_0,
        "ot_hours_3_0": ot_hours_3_0,
        "ot_pay": ot_pay,
        "ot_pay_1_5": ot_pay_1_5,
        "ot_pay_2_0": ot_pay_2_0,
        "ot_pay_3_0": ot_pay_3_0,
        "ot_hourly_rate": ot_hourly_rate,
        "days_employed": days_employed,
        "prorate_factor": round(factor, 4),
        "transport_allowance": round(transport_allowance, 2),
        "meal_allowance": round(meal_allowance, 2),
        "cewi_allowance": round(cewi_allowance, 2),
        "gross_pay": round(gross_pay, 2),
        "epf_employee": round(epf_employee, 2),
        "epf_employer": round(epf_employer, 2),
        "additional_epf_employee": round(additional_epf, 2),
        "socso_employee": round(socso_employee, 2),
        "socso_employer": round(socso_employer, 2),
        "eis_employee": round(eis_employee, 2),
        "eis_employer": round(eis_employer, 2),
        "pcb": round(pcb, 2),
        "pcb_calculated": round(pcb_breakdown["pcb_this_month"], 2),
        "pcb_override": pcb_override,
        "pcb_override_reason": pcb_override_reason,
        "skbbk_employee": round(skbbk_employee, 2),
        "hrd_levy_employer": round(hrd_levy, 2),
        "other_deduction": round(other_deduction, 2),
        "other_deduction_desc": other_deduction_desc,
        "total_deductions": round(total_deductions, 2),
        "net_pay": round(net_pay, 2),
        # ---- Breakdown detail (for the Calculation Detail view) ----
        "calendar_days_in_month": calendar_days_in_month,
        "prorated_basic": round(prorated_basic, 2),
        "basic_salary_raw": basic_salary,
        "meal_eligible_days": meal_eligible_days,
        "cewi_eligible_days": cewi_eligible_days,
        "fixed_allowance_unprorated": round(fixed_allowance, 2),
        "age": age,
        "epf_wage_base": round(epf_wage_base, 2),
        "epf_bracket_label": epf_bracket_label,
        "additional_epf_pct": additional_epf_pct,
        "socso_bracket_label": socso_bracket_label,
        "socso_gate_reason": socso_gate_reason,
        "skbbk_bracket_label": skbbk_bracket_label,
        "skbbk_gate_reason": skbbk_gate_reason,
        "eis_bracket_label": eis_bracket_label,
        "eis_gate_reason": eis_gate_reason,
        "hrd_registered": hrd_registered,
        "hrd_rate": hrd_rate,
        "pcb_breakdown": pcb_breakdown,
    }


def get_payroll_result(conn: sqlite3.Connection, emp_id: str, year: int, month: int):
    """
    Returns the frozen record from a finalized payroll run if one exists for
    this employee/month, otherwise computes a live preview. This is what
    keeps payroll history stable: once a month is finalized, later edits to
    an employee's salary/allowances/attendance must NOT change what already
    shows for that past month.
    """
    conn.row_factory = sqlite3.Row
    row = conn.execute(
        "SELECT * FROM payroll_runs WHERE emp_id=? AND year=? AND month=?",
        (emp_id, year, month),
    ).fetchone()
    if row is None:
        return calculate_payroll(conn, emp_id, year, month)

    result = dict(row)
    emp = conn.execute("SELECT full_name FROM employees WHERE emp_id=?", (emp_id,)).fetchone()
    result["full_name"] = emp["full_name"] if emp else emp_id
    result["ot_hours"] = (result["ot_hours_1_5"] or 0) + (result["ot_hours_2_0"] or 0) + (result["ot_hours_3_0"] or 0)
    return result
