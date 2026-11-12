#!/usr/bin/env python3
"""Build a one-page ATS-safe CV (.docx and .pdf) from a JSON file.

Usage:
  python3 build_cv.py cv.json --layout classic --out my-cv
  (writes my-cv.docx and my-cv.pdf, prints the page count)

Layouts: classic (serif, centred), clean (sans, left aligned), compact (dense, for long histories).
All layouts are single column, no tables except a plain 2-column skills table,
no text boxes, no images, no headers or footers, so ATS parsers read them in order.

JSON shape: see assets/cv.example.json
Exit code 0 = built and one page, 2 = built but more than one page, 1 = error.
"""
import argparse, json, subprocess, sys, os, shutil
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

LAYOUTS = {
    "classic": dict(font="Times New Roman", size=10, name_size=16, align="center",
                    margins=(0.8, 0.6, 1.27, 1.27), heading_caps=False, rule_color="000000",
                    para_after=2, bullet_indent=0.5),
    "clean":   dict(font="Arial", size=9.5, name_size=17, align="left",
                    margins=(1.0, 0.8, 1.5, 1.5), heading_caps=True, rule_color="444444",
                    para_after=2, bullet_indent=0.45),
    "compact": dict(font="Calibri", size=9.5, name_size=15, align="left",
                    margins=(0.7, 0.5, 1.0, 1.0), heading_caps=True, rule_color="777777",
                    para_after=1, bullet_indent=0.4),
}

def set_cell_borders_none(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for b in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{b}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)

def bottom_border(paragraph, color):
    pPr = paragraph._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    b = OxmlElement("w:bottom")
    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "6"); b.set(qn("w:space"), "1"); b.set(qn("w:color"), color)
    pb.append(b); pPr.append(pb)

def run(p, text, L, bold=False, italic=False, size=None):
    r = p.add_run(text)
    r.bold = bold; r.italic = italic
    r.font.name = L["font"]
    r._element.rPr.rFonts.set(qn("w:eastAsia"), L["font"])
    r.font.size = Pt(size or L["size"])
    r.font.color.rgb = RGBColor(0, 0, 0)
    return r

def para(doc, L, align=None, after=None, before=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(L["para_after"] if after is None else after)
    pf.line_spacing = 1.0
    if align == "center": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "justify": p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else: p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p

def heading(doc, L, text):
    p = para(doc, L, align=L["align"], after=2, before=5)
    run(p, text.upper() if L["heading_caps"] else text, L, bold=True, size=L["size"] + (0.5 if not L["heading_caps"] else 0))
    bottom_border(p, L["rule_color"])
    p.paragraph_format.keep_with_next = True

def role_line(doc, L, title, org, place, dates, width_cm):
    p = para(doc, L, after=1, before=2)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(width_cm), WD_TAB_ALIGNMENT.RIGHT)
    run(p, title, L, bold=True)
    if org:
        run(p, " | ", L); run(p, org, L, bold=True)
    if place: run(p, ", " + place, L)
    if dates:
        run(p, "\t", L); run(p, dates, L, italic=True)
    p.paragraph_format.keep_with_next = True

def bullet(doc, L, text):
    p = para(doc, L, align="justify", after=L["para_after"])
    pf = p.paragraph_format
    pf.left_indent = Cm(L["bullet_indent"]); pf.first_line_indent = Cm(-L["bullet_indent"] + 0.05)
    pf.tab_stops.add_tab_stop(Cm(L["bullet_indent"]))
    run(p, "•\t" + text, L)

def build(data, layout, out):
    L = LAYOUTS[layout]
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    top, bottom, left, right = L["margins"]
    sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin = Cm(top), Cm(bottom), Cm(left), Cm(right)
    width_cm = 21.0 - left - right
    st = doc.styles["Normal"]; st.font.name = L["font"]; st.font.size = Pt(L["size"])

    p = para(doc, L, align=L["align"], after=1)
    run(p, data["name"], L, bold=True, size=L["name_size"])
    if data.get("contact"):
        p = para(doc, L, align=L["align"], after=2)
        run(p, "  |  ".join(data["contact"]) if isinstance(data["contact"], list) else data["contact"], L)

    if data.get("profile"):
        heading(doc, L, data.get("profile_heading", "Profile"))
        p = para(doc, L, align="justify"); run(p, data["profile"], L)

    if data.get("experience"):
        heading(doc, L, data.get("experience_heading", "Experience"))
        for j in data["experience"]:
            role_line(doc, L, j["title"], j.get("company", ""), j.get("location", ""), j.get("dates", ""), width_cm)
            for b in j.get("bullets", [])[:3]:
                bullet(doc, L, b)

    for sect in data.get("extra_sections", []):   # e.g. Projects: [{"heading": "Projects", "items": [{"title","org","place","dates","bullets"}]}]
        heading(doc, L, sect["heading"])
        for j in sect["items"]:
            role_line(doc, L, j["title"], j.get("org", ""), j.get("place", ""), j.get("dates", ""), width_cm)
            for b in j.get("bullets", [])[:3]:
                bullet(doc, L, b)

    if data.get("education"):
        heading(doc, L, "Education")
        for e in data["education"]:
            role_line(doc, L, e["degree"], e.get("school", ""), e.get("location", ""), e.get("dates", ""), width_cm)
            if e.get("note"):
                p = para(doc, L); run(p, e["note"], L)

    if data.get("skills"):
        heading(doc, L, "Skills")
        t = doc.add_table(rows=0, cols=2)
        set_cell_borders_none(t)
        t.autofit = False
        lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); t._tbl.tblPr.append(lay)
        grid = t._tbl.tblGrid
        for gc, w in zip(grid.findall(qn("w:gridCol")), (2.6, width_cm - 2.6)):
            gc.set(qn("w:w"), str(int(w / 2.54 * 1440)))
        for label, value in data["skills"]:
            cells = t.add_row().cells
            cells[0].width = Cm(2.6); cells[1].width = Cm(width_cm - 2.6)
            for c, txt, b in ((cells[0], label + ":", True), (cells[1], value, False)):
                p = c.paragraphs[0]; p.paragraph_format.space_after = Pt(1); p.paragraph_format.space_before = Pt(0)
                run(p, txt, L, bold=b)

    docx_path = out + ".docx"
    doc.save(docx_path)
    return docx_path

def to_pdf(docx_path):
    outdir = os.path.dirname(os.path.abspath(docx_path))
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        print("LibreOffice (soffice) not found: the .docx was built but a PDF and page count are not available.")
        return None
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx_path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)
    pdf = os.path.splitext(docx_path)[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None

def pages(pdf):
    try:
        from pypdf import PdfReader
        return len(PdfReader(pdf).pages)
    except Exception:
        r = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True)
        for line in r.stdout.splitlines():
            if line.startswith("Pages:"): return int(line.split()[1])
    return None

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("json_file"); ap.add_argument("--layout", default="classic", choices=LAYOUTS)
    ap.add_argument("--out", required=True, help="output path without extension")
    a = ap.parse_args()
    data = json.load(open(a.json_file, encoding="utf-8"))
    d = build(data, a.layout, a.out)
    pdf = to_pdf(d)
    if pdf:
        n = pages(pdf)
        print(f"built {pdf}: {n} page(s)")
        if n and n > 1:
            print("MORE THAN ONE PAGE: shorten the profile or a bullet, or switch to --layout compact. Do not send this file.")
            sys.exit(2)
