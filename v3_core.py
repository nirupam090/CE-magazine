"""
Core configuration, styles, and reusable helpers for CE Magazine Version 3.
"""

import os, sys, io
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

BASE = r'c:\Users\Nirupam\Documents\magazine ce'
IMG = os.path.join(BASE, 'extracted_images')
OUT = os.path.join(BASE, 'CE_Annual_Magazine_2025-26_v3.docx')

# ── Brand Colors ──
NAVY_HEX = '1B3A5C'
NAVY_DARK_HEX = '0F2440'
GOLD_HEX = 'C8A962'
GOLD_LIGHT_HEX = 'E8DCC0'
GOLD_BG_HEX = 'FBF7EE'
GREEN_HEX = '2E7D32'
WHITE_HEX = 'FFFFFF'
CREAM_HEX = 'FAFAF8'
LIGHT_BG = 'F5F3EE'
ALT_ROW = 'F0EDE6'
BORDER_HEX = 'D8D4CC'
RED_HEX = 'CC3333'
BLUE_BG = 'EEF2F7'
GREEN_BG = 'E8F5E9'
RED_BG = 'FFEBEE'
ORANGE_BG = 'FFF3E0'

NAVY = RGBColor(0x1B, 0x3A, 0x5C)
GOLD = RGBColor(0xC8, 0xA9, 0x62)
GOLD_DARK = RGBColor(0xA0, 0x88, 0x40)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x5A, 0x60, 0x68)
LIGHT_GREY = RGBColor(0x8A, 0x8E, 0x94)

def create_document():
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.8)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    # Base style
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(10)
    style.font.color.rgb = BLACK
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.line_spacing = 1.15

    for level, (sz, sp_b) in {1: (18, 0), 2: (13, 12), 3: (11.5, 8)}.items():
        h = doc.styles[f'Heading {level}']
        h.font.name = 'Cambria'
        h.font.size = Pt(sz)
        h.font.bold = True
        h.font.color.rgb = NAVY
        h.paragraph_format.space_before = Pt(sp_b)
        h.paragraph_format.space_after = Pt(4)
    return doc

def set_cell_shading(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    borders_xml = f'<w:tcBorders {nsdecls("w")}>'
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            color = val.get('color', 'auto')
            sz = val.get('sz', '4')
            b_style = val.get('style', 'single')
            borders_xml += f'<w:{side} w:val="{b_style}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        else:
            borders_xml += f'<w:{side} w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
    borders_xml += '</w:tcBorders>'
    cell._tc.get_or_add_tcPr().append(parse_xml(borders_xml))

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tc_pr = cell._tc.get_or_add_tcPr()
    mar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:start w:w="{left}" w:type="dxa"/>'
        f'  <w:end w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tc_pr.append(mar)

def make_table_full_width(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tw = parse_xml(f'<w:tblW {nsdecls("w")} w:type="pct" w:w="5000"/>')
    existing = tbl_pr.find(qn('w:tblW'))
    if existing is not None:
        tbl_pr.remove(existing)
    tbl_pr.append(tw)

def clear_all_borders(table):
    for row in table.rows:
        for cell in row.cells:
            set_cell_borders(cell)

def run_styled(paragraph, text, font_name='Calibri', size=10, bold=False, italic=False, color=BLACK):
    r = paragraph.add_run(text)
    r.font.name = font_name
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color
    return r

def add_spacer(doc, pts=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(pts)
    p.paragraph_format.line_spacing = Pt(pts)
    return p

def section_header(doc, title, subtitle=None):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, NAVY_DARK_HEX)
    set_cell_margins(cell, top=140, bottom=120, left=200, right=200)
    set_cell_borders(cell, bottom={'color': GOLD_HEX, 'sz': '18', 'style': 'single'})
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if subtitle:
        ps = cell.paragraphs[0]
        ps.paragraph_format.space_after = Pt(1)
        run_styled(ps, subtitle.upper(), 'Calibri', 8.5, color=GOLD, bold=True)
        pt = cell.add_paragraph()
        pt.paragraph_format.space_before = Pt(2)
        pt.paragraph_format.space_after = Pt(2)
        run_styled(pt, title, 'Cambria', 17, bold=True, color=WHITE)
    else:
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        run_styled(p, title, 'Cambria', 17, bold=True, color=WHITE)

    add_spacer(doc, 6)

def sub_section_header(doc, title, num=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    if num:
        run_styled(p, f"{num}  ", 'Cambria', 12.5, bold=True, color=GOLD_DARK)
    run_styled(p, title, 'Cambria', 12.5, bold=True, color=NAVY)
    thin_rule(doc)

def stat_cards(doc, stats):
    n = len(stats)
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)

    for j, (value, label, accent) in enumerate(stats):
        cell = t.cell(0, j)
        accent_hex = {'navy': NAVY_HEX, 'green': GREEN_HEX, 'gold': GOLD_HEX}.get(accent, GOLD_HEX)
        set_cell_shading(cell, LIGHT_BG)
        set_cell_margins(cell, top=110, bottom=90, left=60, right=60)
        set_cell_borders(cell,
            top={'color': accent_hex, 'sz': '16', 'style': 'single'},
            left={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
            right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
            bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        color = {'navy': NAVY, 'green': GREEN, 'gold': GOLD_DARK}.get(accent, NAVY)
        run_styled(p, str(value), 'Cambria', 20, bold=True, color=color)

        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        run_styled(p2, label.upper(), 'Calibri', 7.5, bold=True, color=GREY)

    add_spacer(doc, 6)

def pull_quote(doc, text, attribution=None):
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)

    gc = t.cell(0, 0)
    gc.width = Cm(0.3)
    set_cell_shading(gc, GOLD_HEX)
    set_cell_borders(gc)
    gc.paragraphs[0].text = ''

    qc = t.cell(0, 1)
    set_cell_shading(qc, GOLD_BG_HEX)
    set_cell_margins(qc, top=90, bottom=90, left=150, right=150)
    set_cell_borders(qc)
    p = qc.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run_styled(p, f'“{text}”', 'Cambria', 10.5, italic=True, color=NAVY)

    if attribution:
        pa = qc.add_paragraph()
        pa.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pa.paragraph_format.space_before = Pt(2)
        run_styled(pa, f'— {attribution}', 'Calibri', 8.5, bold=True, color=GOLD_DARK)

    tbl = t._tbl
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if len(cols) >= 2:
            cols[0].set(qn('w:w'), '200')
            cols[1].set(qn('w:w'), '9600')

    add_spacer(doc, 6)

def vision_box(doc, label, text):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, NAVY_DARK_HEX)
    set_cell_margins(cell, top=130, bottom=130, left=180, right=180)
    set_cell_borders(cell, bottom={'color': GOLD_HEX, 'sz': '14', 'style': 'single'})

    pl = cell.paragraphs[0]
    pl.paragraph_format.space_after = Pt(3)
    run_styled(pl, label.upper(), 'Calibri', 8.5, bold=True, color=GOLD)

    pt = cell.add_paragraph()
    pt.paragraph_format.space_before = Pt(2)
    run_styled(pt, f'“{text}”', 'Cambria', 12, italic=True, color=WHITE)

    add_spacer(doc, 6)

def highlight_card(doc, title, text, accent_hex=GOLD_HEX, bg_hex=GOLD_BG_HEX):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, bg_hex)
    set_cell_margins(cell, top=80, bottom=80, left=130, right=130)
    set_cell_borders(cell,
        left={'color': accent_hex, 'sz': '14', 'style': 'single'},
        top={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})
    p = cell.paragraphs[0]
    run_styled(p, title, 'Cambria', 10, bold=True, color=NAVY)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(2)
    run_styled(p2, text, 'Calibri', 9.5, color=BLACK)
    add_spacer(doc, 5)

def data_table(doc, headers, rows, col_widths=None, compact=False):
    sz = 8.5 if compact else 9.5
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)

    # Header
    for j, h in enumerate(headers):
        cell = t.cell(0, j)
        set_cell_shading(cell, NAVY_HEX)
        pad = 60 if compact else 80
        set_cell_margins(cell, top=pad, bottom=pad, left=70, right=70)
        set_cell_borders(cell,
            bottom={'color': GOLD_HEX, 'sz': '12', 'style': 'single'},
            top={'color': NAVY_HEX, 'sz': '4', 'style': 'single'})
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run_styled(p, h.upper(), 'Calibri', sz, bold=True, color=WHITE)

    # Data Rows
    for i, row in enumerate(rows):
        bg = ALT_ROW if i % 2 == 1 else WHITE_HEX
        for j, val in enumerate(row):
            cell = t.cell(i + 1, j)
            set_cell_shading(cell, bg)
            pad = 40 if compact else 55
            set_cell_margins(cell, top=pad, bottom=pad, left=70, right=70)
            set_cell_borders(cell,
                bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
                left={'color': BORDER_HEX, 'sz': '1', 'style': 'single'},
                right={'color': BORDER_HEX, 'sz': '1', 'style': 'single'})
            p = cell.paragraphs[0]
            bold = (j == 0)
            run_styled(p, str(val), 'Calibri', sz, bold=bold, color=BLACK)

    if col_widths:
        for row in t.rows:
            for j, w in enumerate(col_widths):
                if j < len(row.cells):
                    row.cells[j].width = Inches(w)

    add_spacer(doc, 6)

def company_tags(doc, companies):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.3
    for i, c in enumerate(companies):
        run_styled(p, f"  {c}  ", 'Calibri', 8.5, bold=True, color=NAVY)
        if i < len(companies) - 1:
            run_styled(p, " • ", 'Calibri', 8, color=GOLD_DARK)

def gold_rule(doc):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    c = t.cell(0, 0)
    c.height = Pt(2)
    set_cell_shading(c, GOLD_HEX)
    set_cell_borders(c)
    c.paragraphs[0].text = ''
    add_spacer(doc, 4)

def thin_rule(doc):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    c = t.cell(0, 0)
    c.height = Pt(1)
    set_cell_shading(c, BORDER_HEX)
    set_cell_borders(c)
    c.paragraphs[0].text = ''
    add_spacer(doc, 3)

def body(doc, text, bold=False, italic=False, color=BLACK, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    if indent:
        p.paragraph_format.left_indent = Inches(0.2)
    run_styled(p, text, 'Calibri', 10, bold=bold, italic=italic, color=color)
    return p

def body_rich(doc, runs, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    if indent:
        p.paragraph_format.left_indent = Inches(0.2)
    for r_data in runs:
        txt = r_data[0]
        b = r_data[1] if len(r_data) > 1 else False
        it = r_data[2] if len(r_data) > 2 else False
        c = r_data[3] if len(r_data) > 3 else BLACK
        sz = r_data[4] if len(r_data) > 4 else 10
        fn = r_data[5] if len(r_data) > 5 else 'Calibri'
        run_styled(p, txt, fn, sz, bold=b, italic=it, color=c)
    return p

def bullet(doc, text, bold_prefix=None, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
    if bold_prefix:
        run_styled(p, bold_prefix, 'Calibri', 9.8, bold=True, color=NAVY)
    run_styled(p, text, 'Calibri', 9.8, color=BLACK)
    return p

def add_image_safe(doc, path, width=Inches(5.4), caption_text=None, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    if not os.path.exists(path):
        return None
    try:
        p = doc.add_paragraph()
        p.alignment = alignment
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run()
        r.add_picture(path, width=width)
        if caption_text:
            cp = doc.add_paragraph()
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(1)
            cp.paragraph_format.space_after = Pt(6)
            run_styled(cp, caption_text, 'Calibri', 8.5, italic=True, color=GREY)
        return p
    except Exception as e:
        print(f"Error adding image {path}: {e}")
        return None

def caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(5)
    run_styled(p, text, 'Calibri', 8.5, italic=True, color=GREY)

def achievement_card(doc, title, recipient, org, badge=None, badge_color=GOLD_HEX):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    c = t.cell(0, 0)
    set_cell_shading(c, CREAM_HEX)
    set_cell_margins(c, top=70, bottom=70, left=120, right=120)
    set_cell_borders(c,
        left={'color': badge_color, 'sz': '14', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        top={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(1)
    run_styled(p, title, 'Cambria', 10, bold=True, color=NAVY)
    if badge:
        run_styled(p, f"  [{badge}]", 'Calibri', 8, bold=True, color=GOLD_DARK)
    p2 = c.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    run_styled(p2, recipient, 'Calibri', 9, bold=True, color=BLACK)
    run_styled(p2, f"  •  {org}", 'Calibri', 8.5, italic=True, color=GREY)
    add_spacer(doc, 3)

def swoc_box(doc, title, items, border_color=NAVY_HEX, bg_color=LIGHT_BG):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    c = t.cell(0, 0)
    set_cell_shading(c, bg_color)
    set_cell_margins(c, top=80, bottom=80, left=140, right=140)
    set_cell_borders(c,
        left={'color': border_color, 'sz': '16', 'style': 'single'},
        top={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    c_color = NAVY if border_color == NAVY_HEX else (GREEN if border_color == GREEN_HEX else (RED if border_color == RED_HEX else GOLD_DARK))
    run_styled(p, title.upper(), 'Cambria', 10.5, bold=True, color=c_color)
    for it in items:
        pi = c.add_paragraph()
        pi.paragraph_format.space_after = Pt(2)
        pi.paragraph_format.line_spacing = 1.12
        run_styled(pi, "•  ", 'Calibri', 9, bold=True, color=c_color)
        run_styled(pi, it, 'Calibri', 9, color=BLACK)
    add_spacer(doc, 5)
