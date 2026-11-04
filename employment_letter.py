"""Builds a Letter of Employment (.docx) for a hired candidate from the master
Word template in letter_templates/employment_letter.docx.

The template is HR's own sample letter. Fields are found by the wording around
them (not by paragraph number), so HR can edit clauses in the template in Word
and keep things working, as long as these anchor phrases stay:
  - first line of the letter               -> letter date
  - "Name:" ... "Address:"                 -> candidate name / address
  - "I/C Number:"                          -> IC
  - "Dear"                                 -> salutation
  - "RE:" ... "LETTER OF EMPLOYMENT"       -> position
  - "will commence on"                     -> join date
  - "Your monthly remuneration"            -> salary
  - "The official working hours"           -> work schedule
  - "9.2 Service Completion Bonus" .. "RETIREMENT"  -> removed when overseas/bond is off
  - "I," ... "NRIC No"                     -> acknowledgement line
"""
import copy
import datetime
import io
import os
import re

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "letter_templates", "employment_letter.docx")

SCHEDULES = {
    "shift": "Monday to Saturday, on a shift pattern 7am to 3 pm, 3 pm to 11 pm and night 11pm to 7 am, "
             "inclusive of a one-hour lunch break from 12:30 p.m. to 1:30 p.m. In addition, you are required "
             "to work on alternate Saturdays.",
    "office": "Monday to Friday, 8:00 a.m. to 6:00 p.m., inclusive of a one-hour lunch break.",
}

_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven",
         "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _words_under_1000(n):
    parts = []
    if n >= 100:
        parts.append(_ONES[n // 100] + " Hundred")
        n %= 100
    if n >= 20:
        parts.append(_TENS[n // 10] + ("-" + _ONES[n % 10] if n % 10 else ""))
    elif n:
        parts.append(_ONES[n])
    return " ".join(parts)


def ringgit_words(amount):
    """1234.5 -> 'One Thousand Two Hundred Thirty-Four and Cents Fifty Only'."""
    whole = int(amount)
    cents = int(round((amount - whole) * 100))
    if whole == 0:
        words = "Zero"
    else:
        chunks = []
        for value, name in ((1_000_000, "Million"), (1000, "Thousand")):
            if whole >= value:
                chunks.append(_words_under_1000(whole // value) + " " + name)
                whole %= value
        if whole:
            chunks.append(_words_under_1000(whole))
        words = " ".join(chunks)
    if cents:
        words += " and Cents " + _words_under_1000(cents)
    return words + " Only"


def ordinal_date(d):
    """date(2026,12,1) -> '1st December 2026'."""
    day = d.day
    suffix = "th" if 11 <= day % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return f"{day}{suffix} {d.strftime('%B %Y')}"


def format_ic(ic):
    digits = re.sub(r"\D", "", ic or "")
    if len(digits) == 12:
        return f"{digits[:6]}-{digits[6:8]}-{digits[8:]}"
    return (ic or "").strip()


_SKIP_NAMES = {"MUHAMMAD", "MUHAMAD", "MOHAMMAD", "MOHAMAD", "MOHAMED", "MOHD", "MOHD.", "MUHD", "MHD",
               "MOHAMMED", "MUHAMMED", "NUR", "NOR", "NURUL", "SITI", "AL", "AL-", "ABDUL", "ABD", "ABD.", "AHMAD"}


def default_salutation(full_name, salutation=None, gender=None):
    """Best guess for the 'Dear ...' line, e.g. 'Mr Safwan' for MUHAMMAD SAFWAN SAFARDIN.
    HR can overwrite it on the candidate page."""
    title = (salutation or "").strip().rstrip(".")
    if title not in ("Mr", "Ms", "Mrs", "Mdm", "Miss", "Dr"):
        title = "Ms" if (gender or "").lower().startswith("f") else "Mr"
    name = re.split(r"\s+(BIN|BINTI|BT|B\.|A/L|A/P)\s+", (full_name or "").strip().upper())[0]
    words = name.split()
    pick = next((w for w in words if w not in _SKIP_NAMES), words[0] if words else "")
    return f"{title} {pick.title()}".strip()


# ---------- docx helpers ----------

def _set_paragraph(p, segments):
    """Rewrite paragraph p as [(text, bold), ...], keeping the first run's font/formatting."""
    runs = p.runs
    if not runs:
        p.add_run("")
        runs = p.runs
    base = runs[0]._r
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    template = copy.deepcopy(base)
    base.getparent().remove(base)
    from docx.text.run import Run
    for text, bold in segments:
        el = copy.deepcopy(template)
        p._p.append(el)
        run = Run(el, p)
        run.text = text
        run.bold = bold


def _find(paragraphs, pred, start=0):
    for i in range(start, len(paragraphs)):
        if pred(paragraphs[i].text.strip()):
            return i
    return None


def build_letter(data):
    """data keys: full_name, address, ic, salutation_line, position, join_date (date),
    letter_date (date), salary (float), schedule ('shift'|'office'), include_overseas (bool).
    Returns the .docx bytes."""
    import docx

    d = docx.Document(TEMPLATE_PATH)
    P = d.paragraphs
    name = (data["full_name"] or "").strip().upper()
    ic = format_ic(data.get("ic"))
    address = re.sub(r"\s*\n\s*", ", ", (data.get("address") or "").strip()).upper()

    i = _find(P, lambda t: bool(t))
    _set_paragraph(P[i], [(ordinal_date(data["letter_date"]), None)])

    i = _find(P, lambda t: t.startswith("Name:"))
    _set_paragraph(P[i], [("Name:", None), (" " + name, True), ("\n", None),
                          ("Address:", None), (" " + (address or "________________"), True)])

    i = _find(P, lambda t: t.startswith("I/C Number"))
    _set_paragraph(P[i], [("I/C Number:", None), (" " + (ic or "________________"), True)])

    i = _find(P, lambda t: t.startswith("Dear"))
    _set_paragraph(P[i], [("Dear " + data["salutation_line"], None)])

    i = _find(P, lambda t: t.startswith("RE:") and "LETTER OF EMPLOYMENT" in t.upper())
    _set_paragraph(P[i], [("RE:", True), ("\t", True), ("LETTER OF EMPLOYMENT - " + data["position"], True)])

    i = _find(P, lambda t: "will commence on" in t)
    _set_paragraph(P[i], [("Your employment with the Company will commence on ", None),
                          (ordinal_date(data["join_date"]), True),
                          (" subject to the terms and conditions below.", None)])

    i = _find(P, lambda t: t.startswith("Your monthly remuneration"))
    sal = data["salary"]
    _set_paragraph(P[i], [
        ("Your monthly remuneration will be at the rate of ", None),
        (f"Ringgit Malaysia {ringgit_words(sal)} (RM{sal:,.2f})", True),
        (" per month. Payment of salary shall be no later than the seventh (7th) day of the following month. ", None),
        ("Your position is classified as a non-exempt position, and you may be required to work overtime as and "
         "when necessary. Any overtime worked will be compensated in accordance with the Employment Act", True),
        (".", None)])

    i = _find(P, lambda t: t.startswith("The official working hours"))
    _set_paragraph(P[i], [("The official working hours are from ", None),
                          (SCHEDULES.get(data.get("schedule"), SCHEDULES["shift"]), None)])

    if not data.get("include_overseas", True):
        start = _find(P, lambda t: "Service Completion Bonus" in t)
        if start is None:
            start = _find(P, lambda t: t.upper().startswith("OVERSEAS TRAINING"))
        end = _find(P, lambda t: t.upper() == "RETIREMENT", start or 0)
        if start is not None and end is not None:
            for p in P[start:end]:
                p._p.getparent().remove(p._p)
            P = d.paragraphs

    i = _find(P, lambda t: t.startswith("I,") and "NRIC" in t)
    _set_paragraph(P[i], [("I, ", None), (name, None), (", NRIC No: ", None), (ic or "________________", None),
                          (", do hereby agree to, and fully understand and will abide by, the terms and conditions "
                           "as set out in this offer of employment.", None)])

    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()
