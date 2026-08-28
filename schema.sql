-- TMB Payroll App schema. Imported once from Payroll_Master_2026_9.xlsm,
-- then owned by this app going forward.

PRAGMA foreign_keys = ON;

CREATE TABLE employees (
    emp_id                  TEXT PRIMARY KEY,
    full_name               TEXT NOT NULL,
    photo_path               TEXT,          -- stored filename under uploads/<emp_id>/
    ic_passport_no          TEXT,
    date_of_birth           TEXT,          -- ISO date, derived from IC at migration
    race                    TEXT,
    religion                TEXT,
    marital_status          TEXT,
    department              TEXT,
    position                TEXT,
    date_joined             TEXT,          -- ISO date
    status                  TEXT,          -- Active / Resigned / ...
    holiday_state           TEXT,          -- for public holiday matching
    basic_salary            REAL NOT NULL DEFAULT 0,
    working_days_week       REAL,          -- e.g. 5.5 for a Mon-Fri + half-day Saturday schedule
    working_hours_day       REAL,
    standard_start          TEXT,
    standard_end            TEXT,
    lunch_start              TEXT,
    lunch_end                TEXT,
    epf_no                  TEXT,
    socso_no                TEXT,
    tax_no                  TEXT,          -- LHDN income tax reference no.
    eis_no                  TEXT,          -- PERKESO EIS reference no.
    bank_name               TEXT,
    bank_account_no         TEXT,
    phone_number            TEXT,
    hp_no                   TEXT,          -- handphone/mobile number, distinct from Phone Number (office/home line)
    email                   TEXT,
    address                 TEXT,
    emergency_contact_1_name         TEXT,
    emergency_contact_1_phone        TEXT,
    emergency_contact_1_relationship TEXT,
    emergency_contact_2_name         TEXT,
    emergency_contact_2_phone        TEXT,
    emergency_contact_2_relationship TEXT,
    -- Staff self-service portal login. NULL = portal account not yet
    -- activated for this employee (set by HR via Employees > Edit).
    portal_password_hash    TEXT,
    annual_leave_entitlement INTEGER,
    mc_entitlement           INTEGER,
    hospitalisation_leave_entitlement INTEGER,
    medical_claim_limit      REAL DEFAULT 0,   -- RM/year cap for outpatient medical claim reimbursement
    passport_expiry         TEXT,
    work_permit_expiry      TEXT,
    termination_notice_period TEXT,   -- e.g. '1 Month', '2 Months' - notice period required if terminating this employee
    probation_end_date      TEXT,
    -- Set once HR has decided to confirm this employee. When present, the
    -- Alerts page stops listing them as "due" and the confirmation letter
    -- reads its Confirmation Date / New Salary from here by default.
    confirmation_date       TEXT,
    confirmed_new_salary    REAL,
    retirement_date         TEXT,
    last_working_day        TEXT,
    resignation_date        TEXT,  -- date the employee submitted their resignation letter
    skbbk_flag               TEXT,          -- 'Y'/'N' - foreign-worker SKBBK contribution
    work_pattern             TEXT DEFAULT 'Manual', -- 'Manual' / '5-day (Mon-Fri)' /
                                            -- '5.5-day (Mon-Fri + half-day Sat)' / '6-day (Mon-Sat)'
                                            -- - drives auto-fill of Working Days in Month on the
                                            -- Attendance page for employees on a fixed, unambiguous
                                            -- schedule; anyone with an alternating/irregular Saturday
                                            -- pattern must stay 'Manual' since that can't be derived
                                            -- from a calendar alone (confirmed unreliable against a
                                            -- real swipe-card export - see Shamsury/S002 history).
    eis_flag                 TEXT,          -- 'Y'/'N' - manual override on top of the 18-59 age rule
    -- Each allowance below has its own effective date: a payroll run only
    -- applies the allowance for months on/after that date (NULL = no
    -- restriction, applies whenever the Y/N flag is Y). Lets an allowance
    -- change take effect from a specific month without altering past runs.
    transport_allowance      REAL DEFAULT 0,
    transport_allowance_flag TEXT DEFAULT 'N',
    transport_allowance_effective_date TEXT,
    meal_allowance_rate       REAL DEFAULT 0,
    meal_allowance_flag       TEXT DEFAULT 'N',
    meal_allowance_effective_date TEXT,
    position_allowance        REAL DEFAULT 0,
    position_allowance_flag   TEXT DEFAULT 'N',
    position_allowance_effective_date TEXT,
    cewi_rate                 REAL DEFAULT 0,
    cewi_flag                 TEXT DEFAULT 'N',
    cewi_effective_date       TEXT,
    training_incentive        REAL DEFAULT 0,
    training_incentive_flag   TEXT DEFAULT 'N',
    training_incentive_effective_date TEXT,
    oversea_incentive         REAL DEFAULT 0,
    oversea_incentive_flag    TEXT DEFAULT 'N',
    oversea_incentive_effective_date TEXT,
    -- Employee-elected voluntary EPF contribution on top of the statutory
    -- rate, stored as a percentage (e.g. 2 = an extra 2%) of the same EPF
    -- wage base as the statutory contribution. Adds to the employee's EPF
    -- deduction and net-pay reduction; does not change the employer's
    -- statutory contribution.
    additional_epf_employee   REAL DEFAULT 0,
    -- Designated first-line approver (hr_users.username) for this
    -- employee's leave requests, in addition to HR (role='admin'), who can
    -- always approve anyone. NULL = HR only.
    leave_approver_username   TEXT,
    -- Designated supervisor (hr_users.username, role='approver' with
    -- can_approve_appraisal='Y') who appraises this employee - e.g. 'kee'
    -- or 'yang'. NULL = no supervisor assigned yet (won't show up under
    -- anyone's Appraisal team list until set).
    appraisal_supervisor_username TEXT
);

CREATE TABLE public_holidays (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    date    TEXT NOT NULL,   -- ISO date
    day     TEXT,
    name    TEXT NOT NULL,
    remarks TEXT             -- free-text notes (e.g. "Replace of 22/03/2026") -
                              -- not used for any state/coverage filtering
);

-- One row per employee per calendar month: the monthly attendance summary
-- (same shape as each month tab's "MONTHLY SUMMARY" block).
CREATE TABLE attendance_monthly (
    emp_id              TEXT NOT NULL REFERENCES employees(emp_id),
    year                INTEGER NOT NULL,
    month               INTEGER NOT NULL,   -- 1-12
    days_worked         REAL DEFAULT 0,
    al_days             REAL DEFAULT 0,
    mc_days             REAL DEFAULT 0,
    hl_days             REAL DEFAULT 0,
    ul_days             REAL DEFAULT 0,     -- unpaid leave
    other_paid_leave    REAL DEFAULT 0,     -- PL + EL
    ph_days             REAL DEFAULT 0,
    off_days            REAL DEFAULT 0,     -- Saturday rest
    rest_days           REAL DEFAULT 0,     -- Sunday rest
    absent_days         REAL DEFAULT 0,
    working_days_in_month REAL DEFAULT 0,
    late_in_count       INTEGER DEFAULT 0,
    early_out_count     INTEGER DEFAULT 0,
    lunch_late_count    INTEGER DEFAULT 0,
    ot_hours_approved   REAL DEFAULT 0,     -- legacy single-rate total, superseded by the 3 columns below
    -- OT hours split by rate tier (Figure 2.3(a) in the handbook: normal
    -- days and Saturdays are 1.5x, extending to 2.0x late at night;
    -- public holidays are 2.0x, extending to 3.0x late at night). The
    -- payroll preparer assigns each approved OT hour to the tier it
    -- actually falls under; the app just applies the rate.
    ot_hours_1_5        REAL DEFAULT 0,
    ot_hours_2_0        REAL DEFAULT 0,
    ot_hours_3_0        REAL DEFAULT 0,
    meal_eligible_days  REAL DEFAULT 0,
    PRIMARY KEY (emp_id, year, month)
);

CREATE TABLE leave_types (
    code        TEXT PRIMARY KEY,
    description TEXT,
    paid        TEXT,     -- 'Yes' / 'No' / '-'
    notes       TEXT
);

CREATE TABLE leave_entitlement_by_yos (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    min_years_service  REAL NOT NULL,
    annual_leave_days  INTEGER NOT NULL,
    mc_days            INTEGER NOT NULL,
    notes              TEXT
);

-- KWSP/EPF Third Schedule bracket table (verbatim from 'EPF Table')
CREATE TABLE epf_table (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    wage_lower_bound   REAL NOT NULL,
    bracket_label      TEXT,
    part_a_under60_employer REAL,
    part_a_under60_employee REAL,
    part_e_60plus_employer  REAL,
    part_e_60plus_employee  REAL
);

-- PERKESO SOCSO/SKBBK bracket table (verbatim from 'SOCSO-SKBBK Table')
CREATE TABLE socso_table (
    id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    wage_lower_bound       REAL NOT NULL,
    bracket_label          TEXT,
    cat1_employer          REAL,
    cat1_employee_invalidity REAL,
    cat1_employee_skbbk    REAL,
    cat2_employer          REAL,
    cat2_employee_skbbk    REAL
);

-- PERKESO EIS (Employment Insurance System, Act 800) contribution table.
-- Official rates - employer and employee contributions are identical at
-- every bracket (source: perkeso.gov.my, Rate of Contribution Act 800).
CREATE TABLE eis_table (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    wage_lower_bound   REAL NOT NULL,
    bracket_label      TEXT,
    contribution       REAL NOT NULL  -- same amount both employer and employee
);

-- LHDN PCB tax rate schedule (verbatim from 'PCB Tax Rate Table', rows A16:C25)
CREATE TABLE pcb_tax_brackets (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    income_from   REAL NOT NULL,   -- M
    rate          REAL NOT NULL,   -- R
    base_tax      REAL NOT NULL    -- B
);

-- Named constants from 'PCB Tax Rate Table' (B5:B12)
CREATE TABLE pcb_constants (
    key   TEXT PRIMARY KEY,
    value REAL NOT NULL
);

-- Per-employee tax profile, from 'PCB Inputs & YTD' Tax Profile block
CREATE TABLE tax_profile (
    emp_id                TEXT PRIMARY KEY REFERENCES employees(emp_id),
    tax_category           TEXT DEFAULT 'Single',   -- 'Single' / 'Married'
    children_full_relief   INTEGER DEFAULT 0,
    children_half_relief   INTEGER DEFAULT 0,
    tp1_submitted           TEXT DEFAULT '',          -- 'Y' or ''
    tp1_date                TEXT,
    zakat_paid_ytd           REAL DEFAULT 0,
    -- TP3 (Borang PCB/TP3): declared income from a PREVIOUS employer within
    -- the current year, submitted by an employee who joined mid-year. Added
    -- on top of this employer's own YTD figures (pcb_monthly_record) when
    -- computing PCB, so the Computerised Method sees the employee's true
    -- full-year income rather than just what Tianma has paid them.
    tp3_prior_gross          REAL DEFAULT 0,          -- TP3 Part C1: gross remuneration from prior employer(s)
    tp3_prior_epf_employee   REAL DEFAULT 0,          -- TP3 Part C3: employee's EPF contribution at prior employer(s)
    tp3_prior_pcb            REAL DEFAULT 0,          -- TP3 Part C5: PCB already deducted by prior employer(s)
    tp3_submitted            TEXT DEFAULT '',          -- 'Y' or ''
    tp3_date                 TEXT
);

-- Global payroll settings, from 'Payroll' assumptions block (Z1:AA11)
CREATE TABLE payroll_settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- One row per employee per finalized payroll run (parallels 'Payroll History')
CREATE TABLE payroll_runs (
    emp_id              TEXT NOT NULL REFERENCES employees(emp_id),
    year                INTEGER NOT NULL,
    month               INTEGER NOT NULL,
    basic_salary        REAL,
    fixed_allowance     REAL,
    variable_allowance  REAL,
    working_days_in_month REAL,
    days_worked         REAL,
    paid_leave_days     REAL,
    unpaid_days         REAL,
    unpaid_deduction    REAL,
    ot_hours_1_5        REAL,
    ot_hours_2_0        REAL,
    ot_hours_3_0        REAL,
    ot_pay_1_5          REAL,
    ot_pay_2_0          REAL,
    ot_pay_3_0          REAL,
    ot_hourly_rate      REAL,
    ot_pay              REAL,
    days_employed       REAL,
    prorate_factor      REAL,
    transport_allowance REAL,
    meal_allowance      REAL,
    cewi_allowance      REAL,
    gross_pay           REAL,
    epf_employee        REAL,
    additional_epf_employee REAL,
    epf_employer        REAL,
    socso_employee      REAL,
    socso_employer      REAL,
    eis_employee        REAL,
    eis_employer        REAL,
    pcb                 REAL,
    -- Manual correction to the calculated PCB, e.g. from an employee's own
    -- LHDN e-PCB slip when this employer's Computerised Method doesn't
    -- exactly reproduce LHDN's figure (edge cases like irregular voluntary
    -- EPF elections). NULL = use the calculated pcb above; set = pcb,
    -- total_deductions and net_pay all reflect this value instead, and it
    -- survives re-Finalizing (finalize_payroll never clears it).
    pcb_override        REAL,
    pcb_override_reason TEXT,
    skbbk_employee      REAL,
    hrd_levy_employer   REAL,
    other_deduction     REAL,
    other_deduction_desc TEXT,
    total_deductions    REAL,
    net_pay             REAL,
    finalized_at         TEXT,
    PRIMARY KEY (emp_id, year, month)
);

-- Per-employee, per-month ad-hoc Variable Allowance (mirrors 'Payroll'
-- sheet's editable E/AF columns - a manually-entered bonus/allowance that
-- isn't part of the standing employee record and isn't attendance data).
CREATE TABLE monthly_adjustments (
    emp_id              TEXT NOT NULL REFERENCES employees(emp_id),
    year                INTEGER NOT NULL,
    month               INTEGER NOT NULL,
    variable_allowance  REAL DEFAULT 0,
    variable_allowance_flag TEXT DEFAULT 'N',
    -- One-off deduction for that month only (e.g. clawing back an
    -- overpayment) - subtracted from Net Pay only, so it does NOT reduce
    -- Gross Pay or the EPF/SOCSO/EIS/PCB wage base for that month.
    other_deduction     REAL DEFAULT 0,
    other_deduction_desc TEXT,
    PRIMARY KEY (emp_id, year, month)
);

-- Mirrors the workbook's 'PCB Inputs & YTD' Monthly Record log: one row
-- appended per finalized payroll run, used to compute YTD figures for PCB.
CREATE TABLE pcb_monthly_record (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id         TEXT NOT NULL REFERENCES employees(emp_id),
    year           INTEGER NOT NULL,
    month          INTEGER NOT NULL,
    gross_remun    REAL NOT NULL,
    epf_employee   REAL NOT NULL,
    pcb_deducted   REAL NOT NULL,
    UNIQUE(emp_id, year, month)
);

-- Staff self-service leave requests, submitted via the portal and reviewed
-- by HR. Does NOT auto-update attendance_monthly - HR keys the approved
-- days into Attendance separately once a request is approved.
CREATE TABLE leave_requests (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id         TEXT NOT NULL REFERENCES employees(emp_id),
    leave_type     TEXT NOT NULL,   -- matches leave_types.code, e.g. 'AL','MC','UL'
    start_date     TEXT NOT NULL,
    end_date       TEXT NOT NULL,
    days           REAL NOT NULL,
    reason         TEXT,
    status         TEXT NOT NULL DEFAULT 'Pending',  -- Pending / Approved / Rejected
    submitted_at   TEXT NOT NULL,
    reviewed_by    TEXT,
    reviewed_at    TEXT,
    review_notes   TEXT,
    supporting_doc_original TEXT,  -- e.g. medical certificate for MC/Hospitalisation Leave
    supporting_doc_stored   TEXT   -- uploads/<emp_id>/<uuid>_<original filename>
);

-- Business trip / outstation notices - deliberately separate from
-- leave_requests: the employee is still working (just at a different
-- location), so this must never touch leave balance, and doesn't need a
-- supporting document like MC/Hospitalisation Leave does.
CREATE TABLE business_trips (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id         TEXT NOT NULL REFERENCES employees(emp_id),
    destination    TEXT NOT NULL,
    start_date     TEXT NOT NULL,
    end_date       TEXT NOT NULL,
    purpose        TEXT,
    status         TEXT NOT NULL DEFAULT 'Pending',  -- Pending / Approved / Rejected
    submitted_at   TEXT NOT NULL,
    reviewed_by    TEXT,
    reviewed_at    TEXT,
    review_notes   TEXT
);

-- Outpatient medical expense reimbursement claims - tracking/approval only,
-- does NOT flow into payroll automatically (paid out however HR currently
-- handles reimbursements). Balance shown on the Staff Portal is
-- employees.medical_claim_limit minus the sum of this year's Approved
-- claims.
CREATE TABLE medical_claims (
    id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id                 TEXT NOT NULL REFERENCES employees(emp_id),
    claim_date             TEXT NOT NULL,  -- date of treatment
    amount                 REAL NOT NULL,
    clinic_name            TEXT,
    description            TEXT,
    supporting_doc_original TEXT,          -- receipt, as uploaded
    supporting_doc_stored   TEXT,
    status                 TEXT NOT NULL DEFAULT 'Pending',  -- Pending / Approved / Rejected
    submitted_at           TEXT NOT NULL,
    reviewed_by            TEXT,
    reviewed_at            TEXT,
    review_notes           TEXT
);

-- HR/admin accounts that can access the HR side of this app (Employees,
-- Payroll, Attendance, Leave Requests, Business Trips, etc.) - completely
-- separate from employees.portal_password_hash, which only ever grants
-- access to that one employee's own Staff Portal.
-- role='admin' (e.g. Linda/HR): full access to every HR route.
-- role='approver' (e.g. Mr Kee, Mr Yang): restricted by app.py's
-- before_request gate to only the specific sections their can_approve_*
-- flags grant - Leave Requests only for employees whose
-- employees.leave_approver_username matches their own username, and/or
-- Appraisals only for employees whose employees.appraisal_supervisor_username
-- matches their own username. A role='approver' account can have either,
-- both, or (temporarily) neither flag set.
CREATE TABLE hr_users (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    username              TEXT NOT NULL UNIQUE,
    password_hash         TEXT NOT NULL,
    full_name             TEXT NOT NULL,
    created_at            TEXT NOT NULL,
    role                  TEXT NOT NULL DEFAULT 'admin',
    can_approve_leave     TEXT NOT NULL DEFAULT 'N',
    can_approve_appraisal TEXT NOT NULL DEFAULT 'N'
);

-- Employee performance appraisals - digitized version of the paper
-- "Performance Appraisal Form". Scope for now: the supervisor's rating +
-- recommendation only (not the paper form's later Manager/Director
-- countersign, HR verification, or salary recommendation sections - those
-- stay on the printed confirmation_appraisal.html for now). ratings_json
-- holds {"Attitude": 4, "Attendance": 5, ...} for all ~18 factors from the
-- paper form; total_score/percentage are computed at submit time so the
-- HR outcomes list doesn't need to recompute from ratings_json.
CREATE TABLE appraisals (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id                TEXT NOT NULL REFERENCES employees(emp_id),
    purpose               TEXT NOT NULL DEFAULT 'Assessment',  -- Confirmation/Promotion/Increment/Bonus/Assessment
    appraisal_date        TEXT NOT NULL,
    ratings_json          TEXT NOT NULL,   -- {"factor name": 1-5, ...}
    total_score           REAL NOT NULL,   -- sum of ratings
    max_score              REAL NOT NULL,   -- factor count x 5, for percentage calc
    percentage             REAL NOT NULL,   -- total_score / max_score x 100
    overall_band           TEXT NOT NULL,   -- Excellent/Good/Average/Fair/Poor, per the paper form's bands
    rec_advancement        TEXT NOT NULL DEFAULT 'N',  -- (a) qualify for advancement
    rec_not_yet_ready       TEXT NOT NULL DEFAULT 'N',  -- (b) not yet demonstrated required service level
    rec_better_suited       TEXT NOT NULL DEFAULT 'N',  -- (c) better suited for another type of work
    rec_training_required   TEXT NOT NULL DEFAULT 'N',  -- (d) training required
    comments                TEXT,
    -- Salary recommendation, matching the paper form's "Increment" row of
    -- its salary table. current_salary is a snapshot of the employee's
    -- Basic Salary at appraisal time (their real record may change later).
    current_salary          REAL,
    increment_amount        REAL DEFAULT 0,
    new_salary               REAL,
    status                  TEXT NOT NULL DEFAULT 'Draft',  -- Draft (supervisor still editing) / Submitted (final)
    supervisor_username     TEXT NOT NULL,   -- hr_users.username who filled this in
    submitted_at            TEXT,
    hr_viewed_at             TEXT  -- set the first time an HR admin (not the submitting approver) opens it - drives the "new appraisal" nav badge
);

-- Audit trail of probation extensions, separate from just editing
-- employees.probation_end_date directly, so HR can see who pushed the
-- date out, from what, to what, and why. Extending writes a row here and
-- updates employees.probation_end_date to match new_end_date.
CREATE TABLE probation_extensions (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id             TEXT NOT NULL REFERENCES employees(emp_id),
    previous_end_date  TEXT NOT NULL,
    new_end_date       TEXT NOT NULL,
    reason             TEXT,
    extended_at        TEXT NOT NULL
);

-- Audit trail of salary changes, separate from just editing
-- employees.basic_salary directly, so HR can see every past salary, what
-- it changed to, why (Confirmation / Adjustment / Yearly Increment), and
-- when it took effect. Updating writes a row here and updates
-- employees.basic_salary to match new_salary (same no-mid-month-blending
-- behaviour as the rest of payroll in this app).
CREATE TABLE salary_history (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id         TEXT NOT NULL REFERENCES employees(emp_id),
    effective_date TEXT NOT NULL,
    old_salary     REAL NOT NULL,
    new_salary     REAL NOT NULL,
    increment      REAL NOT NULL,
    reason         TEXT NOT NULL,
    recorded_at    TEXT NOT NULL
);

-- Uploaded document files kept per employee (Letter of Employment,
-- Confirmation Letter, Resignation Letter, e-Stamping Certificate, etc.).
-- The actual file bytes live on disk under uploads/<emp_id>/ - this table
-- just tracks what was uploaded, as what type, and when.
CREATE TABLE employee_documents (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id         TEXT NOT NULL REFERENCES employees(emp_id),
    doc_type       TEXT NOT NULL,   -- 'Letter of Employment' / 'Confirmation Letter' / 'Resignation Letter' / 'e-Stamping Certificate' / 'Other'
    original_name  TEXT NOT NULL,
    stored_name    TEXT NOT NULL,   -- filename on disk under uploads/<emp_id>/
    notes          TEXT,
    uploaded_at    TEXT NOT NULL
);
