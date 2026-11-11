"""Renders the Word Letter of Employment (the bytes from employment_letter.build_letter) as an HTML
page (for viewing in the browser) and as a PDF (reportlab), so the View / PDF / Word versions of a
letter always show the same content. Numbering (1., 2.1, (a), bullets) is worked out from the
document's own numbering definitions."""
import base64
import html
import io
import os
import re

import docx
from docx.oxml.ns import qn

LETTERHEAD_PNG = os.path.join(os.path.dirname(__file__), "letter_templates", "letterhead.png")
LETTERHEAD_NAME = "TIANMA PRECISION SDN. BHD."
LETTERHEAD_REG = "202301036641(1530564-X)"
LETTERHEAD_ADDR = "Plo 192, Jalan Bakau 5, Kawasan Perindustrian Tanjung Langsat, 81700 Pasir Gudang, Johor, Malaysia"


def _letters(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(97 + r) + s
    return s


def _numbering_defs(document):
    """{numId: {ilvl: {fmt, text, start, left, hanging}}}"""
    out = {}
    try:
        root = document.part.numbering_part.element
    except Exception:
        return out
    abstracts = {}
    for a in root.findall(qn("w:abstractNum")):
        aid = a.get(qn("w:abstractNumId"))
        lv = {}
        for l in a.findall(qn("w:lvl")):
            ilvl = int(l.get(qn("w:ilvl")))
            fmt = l.find(qn("w:numFmt"))
            txt = l.find(qn("w:lvlText"))
            st = l.find(qn("w:start"))
            ind = l.find(qn("w:pPr") + "/" + qn("w:ind")) if l.find(qn("w:pPr")) is not None else None
            lv[ilvl] = {
                "fmt": fmt.get(qn("w:val")) if fmt is not None else "decimal",
                "text": txt.get(qn("w:val")) if txt is not None else "%1.",
                "start": int(st.get(qn("w:val"))) if st is not None else 1,
                "left": int(ind.get(qn("w:left"), 0)) / 20 if ind is not None else 36,
                "hanging": int(ind.get(qn("w:hanging"), 0)) / 20 if ind is not None else 18,
            }
        abstracts[aid] = lv
    for n in root.findall(qn("w:num")):
        nid = n.get(qn("w:numId"))
        aid = n.find(qn("w:abstractNumId")).get(qn("w:val"))
        out[nid] = abstracts.get(aid, {})
    return out


def _runs_of(p):
    """[(text, bold, underline, italic)] with line breaks as '\n' and tabs as '\t'."""
    runs = []
    for r in p.runs:
        t = r.text
        if t == "":
            continue
        runs.append((t, bool(r.bold), bool(r.underline), bool(r.italic)))
    return runs


def parse(docx_bytes):
    """-> list of blocks. {'kind':'p', runs, align, heading, label, left, hanging}, {'kind':'table', rows},
    {'kind':'rule'}, {'kind':'space'}"""
    d = docx.Document(io.BytesIO(docx_bytes))
    defs = _numbering_defs(d)
    counters = {}
    blocks = []
    tables = {t._tbl: t for t in d.tables}
    from docx.text.paragraph import Paragraph
    for ch in d.element.body.iterchildren():
        if ch.tag == qn("w:tbl"):
            t = tables.get(ch)
            if t is not None:
                rows = [[c.text.strip() for c in r.cells] for r in t.rows]
                plain = all(c.startswith("\u2022") for r in rows for c in r if c)
                blocks.append({"kind": "table", "rows": rows, "plain": plain})
            continue
        if ch.tag != qn("w:p"):
            continue
        p = Paragraph(ch, d)
        text = p.text
        runs = _runs_of(p)
        if not text.strip() and not (runs and any(r[2] for r in runs)):
            blocks.append({"kind": "space"})
            continue
        # a line of dashes followed by a title becomes a rule + the title
        if runs and runs[0][0].strip().startswith("-----"):
            blocks.append({"kind": "rule"})
            runs = runs[1:]
            if not runs:
                continue
        # underlined tabs are blanks to sign / fill in
        if all(r[0].strip("\t ") == "" for r in runs):
            blocks.append({"kind": "p", "runs": [("_" * 34, False, False, False)], "align": "left",
                           "heading": False, "label": None, "left": 0, "hanging": 0})
            continue
        if text.strip() == "Signature\tDate":
            runs = [("Signature: " + "_" * 30 + "          Date: " + "_" * 20, False, False, False)]
        tab_pos = None
        try:
            stops = list(p.paragraph_format.tab_stops)
            if stops and "\t" in text and not text.startswith("Signature"):
                tab_pos = stops[0].position.pt
        except Exception:
            tab_pos = None
        label_runs, value_runs = None, None
        if tab_pos:
            seen = False
            lr, vr = [], []
            for t_, b, u, i in runs:
                if not seen and "\t" in t_ and not (u and t_.strip("\t") == ""):
                    before, _, after = t_.partition("\t")
                    if before:
                        lr.append((before, b, u, i))
                    if after:
                        vr.append((after, b, u, i))
                    seen = True
                elif seen:
                    vr.append((t_, b, u, i))
                else:
                    lr.append((t_, b, u, i))
            if seen and lr:
                label_runs, value_runs = lr, vr
        fixed = []
        for t_, b, u, i in runs:
            if u and t_.strip("\t") == "" and not text.startswith("RE:"):
                t_ = "_" * 22
                u = False
            fixed.append((t_.replace("\t", "    "), b, u, i))
        runs = fixed

        label, left, hanging = None, 0.0, 0.0
        pPr = p._p.pPr
        if pPr is not None and pPr.numPr is not None and pPr.numPr.numId is not None:
            nid = str(pPr.numPr.numId.val)
            ilvl = pPr.numPr.ilvl.val if pPr.numPr.ilvl is not None else 0
            lv = defs.get(nid, {})
            cnt = counters.setdefault(nid, [0] * 9)
            info = lv.get(ilvl, {"fmt": "decimal", "text": "%1.", "start": 1, "left": 36, "hanging": 18})
            if info["fmt"] == "bullet":
                label = "•" if ilvl == 0 else "◦" if ilvl == 1 else "▪"
            else:
                cnt[ilvl] = (cnt[ilvl] + 1) if cnt[ilvl] else info["start"]
                for deeper in range(ilvl + 1, 9):
                    cnt[deeper] = 0
                label = info["text"]
                for k in range(3):
                    lvk = lv.get(k, {"fmt": "decimal"})
                    val = cnt[k]
                    s = _letters(val) if lvk["fmt"] == "lowerLetter" else str(val)
                    label = label.replace(f"%{k + 1}", s)
            left, hanging = info["left"], info["hanging"]
        else:
            li = p.paragraph_format.left_indent
            if li is None and p.style is not None and p.style.paragraph_format.left_indent is not None:
                li = p.style.paragraph_format.left_indent
            left = (li.pt if li is not None else 0)
            ind = p._p.pPr.ind if p._p.pPr is not None else None
            if ind is not None and ind.hanging is not None:
                hanging = ind.hanging.pt
        align = "justify" if p.alignment is not None and int(p.alignment) == 3 else "left"
        heading = bool(p.style is not None and p.style.name.startswith("Heading"))
        _rr = [r for r in runs if r[0].strip()]
        keep_next = heading or (bool(_rr) and len(text.strip()) < 80 and all(r[1] for r in _rr) and not text.strip().endswith("."))             or (text.strip().endswith(":") and len(text.strip()) < 160)
        if heading:
            runs = [(t_, True, u, i) for t_, b, u, i in runs]
        if label_runs is not None and not label:
            x0 = max(left - hanging, 0)
            if heading:
                label_runs = [(t_, True, u, i) for t_, b, u, i in label_runs]
                value_runs = [(t_, True, u, i) for t_, b, u, i in value_runs]
            blocks.append({"kind": "cols", "label": label_runs, "value": value_runs, "x0": x0,
                           "tab": tab_pos, "bold_all": heading, "align": align})
            continue
        blocks.append({"kind": "p", "runs": runs, "align": align, "heading": heading,
                       "label": label, "left": left, "hanging": hanging, "first": -hanging if not label else 0,
                       "keep_next": keep_next})
    # Room to sign: the gap after "Sincerely," (the signing space above the director's name, no line)
    # and the gap above the acceptance "Signature ... Date" line are kept tall instead of collapsing.
    for i, b in enumerate(blocks):
        if b["kind"] != "p":
            continue
        t = "".join(r[0] for r in b["runs"]).strip()
        if t.startswith("Sincerely") and i + 1 < len(blocks) and blocks[i + 1]["kind"] == "space":
            blocks[i + 1]["height"] = 46
        elif t.startswith("Signature:") and i > 0 and blocks[i - 1]["kind"] == "space":
            j = i - 1
            while j > 0 and blocks[j - 1]["kind"] == "space":   # a run of blanks shows as one gap: the first one
                j -= 1
            blocks[j]["height"] = 30
    return blocks


# ---------------- HTML ----------------

def _runs_html(runs):
    out = []
    for t, b, u, i in runs:
        s = html.escape(t).replace("\n", "<br>")
        if b:
            s = f"<b>{s}</b>"
        if u:
            s = f"<u>{s}</u>"
        if i:
            s = f"<i>{s}</i>"
        out.append(s)
    return "".join(out)


def _letterhead_b64():
    with open(LETTERHEAD_PNG, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def render_html(docx_bytes, title, pdf_url, word_url, back_url):
    parts = []
    prev_space = True
    for b in parse(docx_bytes):
        if b["kind"] == "space":
            if b.get("height") and not prev_space:
                parts.append(f'<div class="gap" style="height:{int(b["height"] * 1.33)}px"></div>')
            elif not prev_space:
                parts.append('<div class="gap"></div>')
            prev_space = True
            continue
        prev_space = False
        if b["kind"] == "rule":
            parts.append("<hr>")
        elif b["kind"] == "cols":
            w = max(b["tab"] - b["x0"], 20)
            cls = "p h" if b["bold_all"] else "p"
            pre = f"margin-left:{b['x0']:.0f}px;"
            parts.append(f'<div class="{cls}" style="{pre} display:flex; text-align:{b["align"]};">'
                         f'<span style="flex:0 0 {w:.0f}px;">{_runs_html(b["label"])}</span>'
                         f'<span style="flex:1;">{_runs_html(b["value"])}</span></div>')
        elif b["kind"] == "table":
            if b.get("plain"):
                rows = "".join("<tr>" + "".join(f"<td>{html.escape(c)}</td>" for c in r) + "</tr>" for r in b["rows"])
                parts.append(f'<table class="plain">{rows}</table>')
            else:
                rows = "".join(
                    "<tr>" + "".join(f"<{'th' if ri == 0 else 'td'}>{html.escape(c)}</{'th' if ri == 0 else 'td'}>" for c in r) + "</tr>"
                    for ri, r in enumerate(b["rows"]))
                parts.append(f'<table class="grid">{rows}</table>')
        else:
            cls = "h" if b["heading"] else ""
            left = max(b["left"], 0)
            if b["label"]:
                hang = max(b["hanging"], 14)
                parts.append(
                    f'<div class="p {cls}" style="margin-left:{left:.0f}px; padding-left:0; text-align:{b["align"]}; display:flex;">'
                    f'<span style="flex:0 0 {hang:.0f}px; margin-left:-{hang:.0f}px; text-align:left;">{html.escape(b["label"])}</span>'
                    f'<span style="flex:1;">{_runs_html(b["runs"])}</span></div>')
            else:
                first = b.get("first", 0)
                ti = f" text-indent:{first:.0f}px;" if first else ""
                parts.append(f'<div class="p {cls}" style="margin-left:{left:.0f}px;{ti} text-align:{b["align"]};">{_runs_html(b["runs"])}</div>')
    body = "\n".join(parts)
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<style>
  body {{ margin:0; background:#e5e7eb; font-family: Arial, Helvetica, sans-serif; color:#000; }}
  .bar {{ position:sticky; top:0; z-index:5; background:#1e293b; padding:8px 16px; display:flex; gap:8px; flex-wrap:wrap; align-items:center; }}
  .bar a, .bar button {{ background:#fff; color:#111827; border:0; border-radius:6px; padding:6px 14px; font-size:14px; text-decoration:none; cursor:pointer; }}
  .bar a.primary {{ background:#2563eb; color:#fff; }}
  .bar span {{ color:#cbd5e1; font-size:13px; margin-left:auto; }}
  .page {{ width:794px; max-width:calc(100% - 24px); box-sizing:border-box; margin:16px auto; background:#fff; padding:40px 56px 56px 72px; box-shadow:0 2px 12px rgba(0,0,0,.2); font-size:13.3px; line-height:1.4; }}
  .lh {{ margin-bottom:16px; }}
  .p {{ margin:0 0 6px; }}
  .p.h {{ margin-top:8px; }}
  .gap {{ height:8px; }}
  hr {{ border:0; border-top:1px solid #000; margin:14px 0 10px; }}
  table.grid {{ border-collapse:collapse; margin:4px 0 10px 28px; min-width:60%; }}
  table.grid th, table.grid td {{ border:1px solid #444; padding:4px 10px; text-align:left; }}
  table.grid th {{ background:#f1f5f9; }}
  table.plain {{ border-collapse:collapse; margin:0 0 8px 60px; width:calc(100% - 60px); }}
  table.plain td {{ padding:1px 8px 1px 0; width:50%; vertical-align:top; }}
  .lh img {{ display:block; width:100%; height:auto; }}
  @media print {{
    body {{ background:#fff; }} .bar {{ display:none; }}
    .page {{ box-shadow:none; margin:0; width:auto; max-width:none; padding:0; }}
    @page {{ size:A4; margin:14mm 12mm 16mm 18mm; }}
  }}
</style></head><body>
<div class="bar">
  <a href="{html.escape(back_url)}">&laquo; Back to candidate</a>
  <button type="button" onclick="window.print()">Print</button>
  <a class="primary" href="{html.escape(pdf_url)}">Download PDF</a>
  <a href="{html.escape(word_url)}">Download Word</a>
  <span>Preview &mdash; edit the details on the candidate page, then view again.</span>
</div>
<div class="page">
  <div class="lh"><img alt="Tianma Precision Sdn. Bhd." src="data:image/png;base64,{_letterhead_b64()}"></div>
{body}
</div></body></html>"""


# ---------------- PDF ----------------

def render_pdf(docx_bytes):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.platypus import HRFlowable, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    class Numbered(pdfcanvas.Canvas):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            self._saved = []

        def showPage(self):
            self._saved.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._saved)
            for st in self._saved:
                self.__dict__.update(st)
                self.setFont("Helvetica", 8)
                self.setFillColor(colors.HexColor("#444444"))
                self.drawCentredString(67 + (A4[0] - 67 - 46) / 2, 34, f"Page {self._pageNumber} of {total}")
                super().showPage()
            super().save()

    def header(c, doc):
        # the company letterhead exactly as in the Word template (image of its header)
        c.saveState()
        w_pt = A4[0] - 67 - 46
        img_h = w_pt * 183 / 1931
        c.drawImage(LETTERHEAD_PNG, 67, A4[1] - 22 - img_h, width=w_pt, height=img_h)
        c.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=67, rightMargin=46, topMargin=86, bottomMargin=56,
                            title="Letter of Employment", author="Tianma Precision Sdn. Bhd.")

    def esc(t):
        return html.escape(t).replace("\n", "<br/>")

    def runs_markup(runs):
        out = []
        for t, b, u, i in runs:
            s = esc(t)
            if b:
                s = f"<b>{s}</b>"
            if u:
                s = f"<u>{s}</u>"
            if i:
                s = f"<i>{s}</i>"
            out.append(s)
        return "".join(out)

    story = []
    prev_space = True
    avail_w = A4[0] - 67 - 46
    for b in parse(docx_bytes):
        if b["kind"] == "space":
            if not prev_space:
                story.append(Spacer(1, b.get("height") or 6))
            prev_space = True
            continue
        prev_space = False
        if b["kind"] == "cols":
            w = max(b["tab"] - b["x0"], 20)
            base = ParagraphStyle("cl", fontName="Helvetica", fontSize=9.6, leading=12.6,
                                  alignment=TA_JUSTIFY if b["align"] == "justify" else TA_LEFT)
            lab = Paragraph(runs_markup(b["label"]), base)
            val = Paragraph(runs_markup(b["value"]), base)
            tbl = Table([[lab, val]], colWidths=[w, avail_w - b["x0"] - w], hAlign="LEFT")
            tbl.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                     ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
                                     ("VALIGN", (0, 0), (-1, -1), "TOP")]))
            wrapper = Table([[tbl]], colWidths=[avail_w], hAlign="LEFT")
            wrapper.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), b["x0"]), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                         ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
            story.append(wrapper)
        elif b["kind"] == "rule":
            story.append(Spacer(1, 6))
            story.append(HRFlowable(width="100%", thickness=0.8, color=colors.black))
            story.append(Spacer(1, 6))
        elif b["kind"] == "table":
            plain = b.get("plain")
            data = [[Paragraph(esc(c), ParagraphStyle("c", fontName="Helvetica-Bold" if (ri == 0 and not plain) else "Helvetica",
                                                      fontSize=9.6 if plain else 9.5, leading=12.6 if plain else 12))
                     for c in r] for ri, r in enumerate(b["rows"])]
            ncol = max(len(r) for r in data)
            if plain:
                t = Table(data, colWidths=[(avail_w - 41) / ncol] * ncol, hAlign="RIGHT")
                t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 1),
                                       ("BOTTOMPADDING", (0, 0), (-1, -1), 1), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
            else:
                t = Table(data, colWidths=[(avail_w - 60) / ncol] * ncol, hAlign="LEFT")
                t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#444444")),
                                       ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f1f5f9")),
                                       ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
            story.append(Spacer(1, 2))
            story.append(t)
            story.append(Spacer(1, 8))
        else:
            left = max(b["left"], 0)
            hang = max(b["hanging"], 14) if b["label"] else 0
            style = ParagraphStyle(
                "p", fontName="Helvetica", fontSize=9.6, leading=12.6, spaceAfter=4.5,
                spaceBefore=6 if b["heading"] else 0,
                alignment=TA_JUSTIFY if b["align"] == "justify" else TA_LEFT,
                leftIndent=left, firstLineIndent=b.get("first", 0), keepWithNext=1 if (b["heading"] or b.get("keep_next")) else 0)
            markup = runs_markup(b["runs"])
            if b["label"]:
                style.leftIndent = left
                style.firstLineIndent = -hang
                markup = f'{esc(b["label"])}  ' + markup
                # tab-like alignment: place the text after the hanging gap
                lab_w = hang
                markup = (f'<font name="Helvetica">{esc(b["label"])}</font>')
                para = Paragraph(markup, ParagraphStyle("lab", fontName="Helvetica", fontSize=9.6, leading=12.6))
                body = Paragraph(runs_markup(b["runs"]), ParagraphStyle(
                    "pb", parent=style, leftIndent=0, firstLineIndent=0, spaceAfter=0, spaceBefore=0))
                tbl = Table([[para, body]], colWidths=[hang, avail_w - left], hAlign="LEFT")
                tbl.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                         ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
                                         ("VALIGN", (0, 0), (-1, -1), "TOP")]))
                # indent the whole row so the label hangs into the margin
                wrapper = Table([[tbl]], colWidths=[avail_w], hAlign="LEFT")
                wrapper.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), max(left - hang, 0)),
                                             ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                             ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
                wrapper.keepWithNext = 1 if (b["heading"] or b.get("keep_next")) else 0
                story.append(wrapper)
            else:
                story.append(Paragraph(markup, style))
    doc.build(story, onFirstPage=header, onLaterPages=header, canvasmaker=Numbered)
    return buf.getvalue()
