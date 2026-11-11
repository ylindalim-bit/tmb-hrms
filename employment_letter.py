"""Builds a Letter of Employment (.docx) for a hired candidate from the master
Word template in letter_templates/employment_letter.docx.

The template is HR's own sample letter. Fields are found by the wording around
them (not by paragraph number), so HR can edit clauses in the template in Word
and keep things working, as long as these anchor phrases stay:
  - first line of the letter               -> letter date
  - "Name"                                 -> candidate name (an Address line is added after it)
  - "I/C Number"                           -> IC
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


def default_salutation(full_name=None, salutation=None, gender=None):
    """Default for the 'Dear ...' line: 'Sir', or 'Madam' for a female candidate.
    HR can overwrite it on the candidate page."""
    title = (salutation or "").strip().rstrip(".")
    female = title in ("Ms", "Mrs", "Mdm", "Miss") or (gender or "").lower().startswith("f")
    return "Madam" if female else "Sir"


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


def _insert_after(p, segments):
    """New paragraph right after p, same paragraph style/format, filled with segments."""
    new = copy.deepcopy(p._p)
    p._p.addnext(new)
    from docx.text.paragraph import Paragraph
    np_ = Paragraph(new, p._parent)
    _set_paragraph(np_, segments)
    return np_


def _label_layout(p):
    """'Label<TAB>: value' with the value column lined up and long values wrapping under it."""
    from docx.shared import Inches
    pf = p.paragraph_format
    pf.left_indent = Inches(1.15)
    pf.first_line_indent = Inches(-1.15)
    pf.space_after = 0
    pf.tab_stops.add_tab_stop(Inches(1.0))


def _find(paragraphs, pred, start=0):
    for i in range(start, len(paragraphs)):
        if pred(paragraphs[i].text.strip()):
            return i
    return None


def _holidays_to_table(d):
    """The Public Holidays bullet list becomes a 2-column table (list order runs down the left column,
    then down the right), keeping each holiday's own formatting."""
    import docx
    from docx.oxml.ns import qn
    from docx.shared import Twips
    P = d.paragraphs
    first = _find(P, lambda t: t.startswith("New Year"))
    if first is None:
        return

    def num_id(par):
        pPr = par._p.pPr
        if pPr is None or pPr.numPr is None or pPr.numPr.numId is None:
            return None
        return str(pPr.numPr.numId.val)

    nid = num_id(P[first])
    if nid is None:
        return
    items = []
    i = first
    while i < len(P) and num_id(P[i]) == nid and P[i].text.strip():
        items.append(P[i])
        i += 1
    if len(items) < 4:
        return
    rows = (len(items) + 1) // 2
    tbl = d.add_table(rows=rows, cols=2)
    tbl.autofit = False
    tblPr = tbl._tbl.tblPr
    # no borders, table indented to line up with the clause text
    ind = tblPr.makeelement(qn("w:tblInd"), {qn("w:w"): "820", qn("w:type"): "dxa"})
    anchor = tblPr.find(qn("w:tblLayout"))
    if anchor is None:
        anchor = tblPr.find(qn("w:tblLook"))
    if anchor is not None:
        anchor.addprevious(ind)      # keep the schema order Word expects
    else:
        tblPr.append(ind)
    for col, width in zip(tbl.columns, (4300, 4300)):
        col.width = Twips(width)
    for k, src in enumerate(items):
        r, c = (k, 0) if k < rows else (k - rows, 1)
        cell = tbl.cell(r, c)
        cell.width = Twips(4300)
        new_p = copy.deepcopy(src._p)
        pPr = new_p.pPr
        if pPr is not None:
            for tag in ("w:numPr", "w:ind"):
                el = pPr.find(qn(tag))
                if el is not None:
                    pPr.remove(el)
        # a bullet character in front of the first run
        runs = new_p.findall(qn("w:r"))
        if runs:
            bullet = copy.deepcopy(runs[0])
            for t in bullet.findall(qn("w:t")):
                bullet.remove(t)
            t = bullet.makeelement(qn("w:t"), {})
            t.text = "\u2022\u00a0\u00a0"
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            bullet.append(t)
            runs[0].addprevious(bullet)
        old_p = cell.paragraphs[0]._p
        old_p.addprevious(new_p)
        old_p.getparent().remove(old_p)
    items[0]._p.addprevious(tbl._tbl)
    for src in items:
        src._p.getparent().remove(src._p)


def _keep_titles_with_text(d):
    """A clause title (SALARY, Transport Allowance, 9.1 Performance Bonus ...) must never be left alone at the
    bottom of a page: mark every short all-bold title, and every lead-in line ending with ':', 'keep with next'
    so Word moves the title to the next page together with its text."""
    from docx.oxml.ns import qn
    for p in d.paragraphs:
        t = p.text.strip()
        runs = [r for r in p.runs if r.text.strip()]
        if not t or not runs:
            continue
        is_title = len(t) < 80 and all(r.bold for r in runs) and not t.endswith(".")
        is_lead_in = t.endswith(":") and len(t) < 160
        if is_title or is_lead_in or p.style.name.startswith("Heading"):
            p.paragraph_format.keep_with_next = True
    for tbl in d.tables:
        for row in tbl.rows:
            trPr = row._tr.get_or_add_trPr()
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(trPr.makeelement(qn("w:cantSplit"), {}))


def set_ref_on_letter(docx_bytes, ref):
    """An uploaded (hand-made) letter, used as it is, except that its Ref line is set: an existing 'Ref:' line is
    updated, otherwise a 'Ref: ...' line is added under the date line."""
    import docx
    d = docx.Document(io.BytesIO(docx_bytes))
    P = d.paragraphs
    if ref:
        i = _find(P, lambda t: t.lower().startswith("ref:") or t.lower().startswith("ref :"))
        if i is not None:
            _set_paragraph(P[i], [("Ref: " + ref, None)])
        else:
            j = _find(P, lambda t: bool(t))
            if j is not None:
                _insert_after(P[j], [("Ref: " + ref, None)])
    _keep_titles_with_text(d)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def build_letter(data):
    """data keys: ref (optional, e.g. 'TMB/HR/LOE/2026/011'), full_name, address, ic, salutation_line, position, join_date (date),
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
    if data.get("ref"):
        _insert_after(P[i], [("Ref: " + data["ref"], None)])
        P = d.paragraphs

    i = _find(P, lambda t: t.startswith("Name:"))
    _set_paragraph(P[i], [("Name", None), ("	: ", None), (name, True)])
    _label_layout(P[i])
    addr = _insert_after(P[i], [("Address", None), ("	: ", None), (address or "________________", True)])
    _label_layout(addr)
    P = d.paragraphs

    i = _find(P, lambda t: t.startswith("I/C Number"))
    _set_paragraph(P[i], [("I/C Number", None), ("	: ", None), (ic or "________________", True)])
    _label_layout(P[i])

    i = _find(P, lambda t: t.startswith("Dear"))
    greeting = re.sub(r"^\s*dear\b\s*", "", data.get("salutation_line") or "", flags=re.I).strip().rstrip(",")
    greeting = (greeting[:1].upper() + greeting[1:]) if greeting else "Sir"
    _set_paragraph(P[i], [("Dear " + greeting + ",", None)])

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

    _holidays_to_table(d)
    P = d.paragraphs

    i = _find(P, lambda t: t.startswith("I,") and "NRIC" in t)
    _set_paragraph(P[i], [("I, ", None), (name, None), (", NRIC No: ", None), (ic or "________________", None),
                          (", do hereby agree to, and fully understand and will abide by, the terms and conditions "
                           "as set out in this offer of employment.", None)])

    _keep_titles_with_text(d)

    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()
