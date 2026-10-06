"""
CE Department Annual Magazine 2025-26 — MASTER DOCX Generator (v3 - Audit & CSE Reference Updated)
Bharatiya Vidya Bhavan's Sardar Patel Institute of Technology, Mumbai

Features:
- Full alignment with '02_10_2026 External Audit Computer Engineering Department 2024-25 and 2025-26_UPDATED.pdf'
- Authentic Principal's Desk address from 'CSE Magazine AY-2025-2026.pdf' with real photo
- Hierarchical Table of Contents with numbered headings, subheadings, and indentation matching 'Copy of Magazine INDEX.docx'
- Embedded high-res charts: Research funding (₹6.05 Cr), Placement trends & tiers, Research mix (43 papers), Faculty development (8,600+ hrs), Student outcomes
- Embedded high-res photographs: Principal Dr. Chaudhari, Faculty & Staff, PTM, IEEE CS events, and Student Hackathon / GATE Victories
- Professional typography, Navy & Gold institutional styling, pull quotes, stat cards, and data tables
"""

import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from PIL import Image
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

BASE = r'c:\Users\Nirupam\Documents\magazine ce'
IMG = os.path.join(BASE, 'extracted_images')
OUT = os.path.join(BASE, 'CE_Annual_Magazine_2025-26_v2.docx')

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
BLUE_SOFT = 'E3EDF7'

NAVY = RGBColor(0x1B, 0x3A, 0x5C)
GOLD = RGBColor(0xC8, 0xA9, 0x62)
GOLD_DARK = RGBColor(0xA0, 0x88, 0x40)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x5A, 0x60, 0x68)
LIGHT_GREY = RGBColor(0x8A, 0x8E, 0x94)

doc = Document()

# ═══════════════════════════════════════
# PAGE SETUP (A4 Portrait)
# ═══════════════════════════════════════
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(2.0)
sec.right_margin = Cm(2.0)

# Default Normal Style
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
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


# ═══════════════════════════════════════
# UTILITIES & DESIGN SYSTEM
# ═══════════════════════════════════════

def set_cell_shading(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    borders_xml = f'<w:tcBorders {nsdecls("w")}>'
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            color = val.get('color', 'auto')
            sz = val.get('sz', '4')
            style = val.get('style', 'single')
            borders_xml += f'<w:{side} w:val="{style}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
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

def run_styled(paragraph, text, font_name='Calibri', size=10.5, bold=False, italic=False, color=BLACK):
    r = paragraph.add_run(text)
    r.font.name = font_name
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.font.color.rgb = color
    return r

def add_spacer(pts=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(pts)
    p.paragraph_format.line_spacing = Pt(pts)
    return p

def section_header(title, subtitle=None):
    """Full-width navy banner with gold accent bar."""
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
        run_styled(pt, title, 'Cambria', 18, bold=True, color=WHITE)
    else:
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        run_styled(p, title, 'Cambria', 18, bold=True, color=WHITE)

    add_spacer(8)

def sub_section_header(title, num=None):
    """Sub-heading bar with subtle gold/navy accent line."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    if num:
        run_styled(p, f"{num}  ", 'Cambria', 12.5, bold=True, color=GOLD_DARK)
    run_styled(p, title, 'Cambria', 12.5, bold=True, color=NAVY)
    thin_rule()

def stat_cards(stats):
    """Row of modern KPI metric boxes."""
    n = len(stats)
    t = doc.add_table(rows=1, cols=n)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)

    for j, (value, label, accent) in enumerate(stats):
        cell = t.cell(0, j)
        accent_hex = {'navy': NAVY_HEX, 'green': GREEN_HEX, 'gold': GOLD_HEX}.get(accent, GOLD_HEX)
        set_cell_shading(cell, LIGHT_BG)
        set_cell_margins(cell, top=120, bottom=100, left=60, right=60)
        set_cell_borders(cell,
            top={'color': accent_hex, 'sz': '16', 'style': 'single'},
            left={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
            right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
            bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})

        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        color = {'navy': NAVY, 'green': GREEN, 'gold': GOLD_DARK}.get(accent, NAVY)
        run_styled(p, str(value), 'Cambria', 22, bold=True, color=color)

        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        run_styled(p2, label.upper(), 'Calibri', 7.5, bold=True, color=GREY)

    add_spacer(8)

def pull_quote(text, attribution=None):
    """Editorial quote box with vertical gold border."""
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
    set_cell_margins(qc, top=100, bottom=100, left=160, right=160)
    set_cell_borders(qc)
    p = qc.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run_styled(p, f'“{text}”', 'Cambria', 11, italic=True, color=NAVY)

    if attribution:
        pa = qc.add_paragraph()
        pa.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pa.paragraph_format.space_before = Pt(2)
        run_styled(pa, f'— {attribution}', 'Calibri', 9, bold=True, color=GOLD_DARK)

    # Column proportions
    tbl = t._tbl
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if len(cols) >= 2:
            cols[0].set(qn('w:w'), '200')
            cols[1].set(qn('w:w'), '9600')

    add_spacer(6)

def vision_box(label, text):
    """Executive dark navy banner for institutional vision."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, NAVY_DARK_HEX)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    set_cell_borders(cell, bottom={'color': GOLD_HEX, 'sz': '14', 'style': 'single'})

    pl = cell.paragraphs[0]
    pl.paragraph_format.space_after = Pt(4)
    run_styled(pl, label.upper(), 'Calibri', 8.5, bold=True, color=GOLD)

    pt = cell.add_paragraph()
    pt.paragraph_format.space_before = Pt(2)
    run_styled(pt, f'“{text}”', 'Cambria', 12.5, italic=True, color=WHITE)

    add_spacer(8)

def highlight_card(title, text, accent_hex=GOLD_HEX, bg_hex=GOLD_BG_HEX):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, bg_hex)
    set_cell_margins(cell, top=90, bottom=90, left=140, right=140)
    set_cell_borders(cell,
        left={'color': accent_hex, 'sz': '14', 'style': 'single'},
        top={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})
    p = cell.paragraphs[0]
    run_styled(p, title, 'Cambria', 10.5, bold=True, color=NAVY)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_before = Pt(2)
    run_styled(p2, text, 'Calibri', 9.5, color=BLACK)
    add_spacer(6)

def data_table(headers, rows, col_widths=None, compact=False):
    """Executive table with navy header band and alternating shading."""
    sz = 8.5 if compact else 9.5
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    t.style = 'Table Grid'

    for j, h in enumerate(headers):
        c = t.cell(0, j)
        set_cell_shading(c, NAVY_HEX)
        set_cell_margins(c, top=45, bottom=45, left=70, right=70)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run_styled(p, h, 'Calibri', sz - 0.5, bold=True, color=WHITE)

    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i + 1, j)
            if i % 2 == 1:
                set_cell_shading(c, ALT_ROW)
            set_cell_margins(c, top=30, bottom=30, left=70, right=70)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(1.5)
            p.paragraph_format.space_after = Pt(1.5)
            is_bold = (j == 0 and len(headers) > 2) or str(val).startswith('*')
            val_text = str(val).lstrip('*')
            run_styled(p, val_text, 'Calibri', sz, bold=is_bold)

    if col_widths:
        for j, w in enumerate(col_widths):
            for row_obj in t.rows:
                row_obj.cells[j].width = Cm(w)

    add_spacer(6)
    return t

def gold_rule():
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_borders(cell, bottom={'color': GOLD_HEX, 'sz': '8', 'style': 'single'})
    cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    cell.paragraphs[0].paragraph_format.line_spacing = Pt(1)
    set_cell_margins(cell, 0, 0, 0, 0)
    add_spacer(6)

def thin_rule():
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_borders(cell, bottom={'color': BORDER_HEX, 'sz': '4', 'style': 'single'})
    cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    cell.paragraphs[0].paragraph_format.line_spacing = Pt(1)
    set_cell_margins(cell, 0, 0, 0, 0)
    add_spacer(4)

def body(text, indent=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.7)
    run_styled(p, text, 'Calibri', 10.5, italic=italic)
    return p

def body_rich(segments, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    for seg in segments:
        txt = seg[0]
        bold = seg[1] if len(seg) > 1 else False
        ital = seg[2] if len(seg) > 2 else False
        clr = seg[3] if len(seg) > 3 else BLACK
        run_styled(p, txt, 'Calibri', 10.5, bold=bold, italic=ital, color=clr)
    return p

def bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    if bold_prefix:
        run_styled(p, bold_prefix, 'Calibri', 10, bold=True, color=NAVY)
        run_styled(p, text, 'Calibri', 10)
    else:
        run_styled(p, text, 'Calibri', 10)

def placeholder(text):
    p = doc.add_paragraph()
    r = p.add_run(f'  ▶ {text}  ')
    r.font.name = 'Calibri'
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0xCC, 0x33, 0x33)
    r.bold = True
    rPr = r._r.get_or_add_rPr()
    rPr.append(parse_xml(f'<w:highlight {nsdecls("w")} w:val="yellow"/>'))

def add_image(path, width=Inches(5.4), caption_text=None):
    full = path if os.path.isabs(path) else os.path.join(BASE, path)
    if os.path.exists(full):
        try:
            doc.add_picture(full, width=width)
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if caption_text:
                caption(caption_text)
            else:
                add_spacer(4)
            return True
        except Exception as e:
            placeholder(f'IMAGE FORMAT ERROR: {os.path.basename(path)}')
            return False
    else:
        placeholder(f'IMAGE NOT FOUND: {os.path.basename(path)}')
        return False

def add_two_images(img1_path, cap1, img2_path, cap2, width=Inches(2.7)):
    """Place two images neatly side-by-side in a 2-column table."""
    full1 = img1_path if os.path.isabs(img1_path) else os.path.join(BASE, img1_path)
    full2 = img2_path if os.path.isabs(img2_path) else os.path.join(BASE, img2_path)
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    clear_all_borders(t)

    for idx, (fpath, cap) in enumerate([(full1, cap1), (full2, cap2)]):
        cell = t.cell(0, idx)
        set_cell_margins(cell, top=40, bottom=40, left=40, right=40)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(fpath):
            try:
                r = p.add_run()
                r.add_picture(fpath, width=width)
                if cap:
                    pc = cell.add_paragraph()
                    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    pc.paragraph_format.space_before = Pt(2)
                    run_styled(pc, cap, 'Calibri', 8, italic=True, color=GREY)
            except:
                p.text = f'[Image: {os.path.basename(fpath)}]'
        else:
            p.text = f'[Image not found: {os.path.basename(fpath)}]'
    add_spacer(6)

def caption(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    run_styled(p, text, 'Calibri', 8.5, italic=True, color=GREY)

def achievement_card(event, students, org, result, accent=GOLD_HEX):
    t = doc.add_table(rows=1, cols=1)
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, LIGHT_BG)
    set_cell_margins(cell, top=80, bottom=80, left=160, right=120)
    set_cell_borders(cell,
        left={'color': accent, 'sz': '16', 'style': 'single'},
        top={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})

    p1 = cell.paragraphs[0]
    p1.paragraph_format.space_after = Pt(1)
    run_styled(p1, event, 'Cambria', 10.5, bold=True, color=NAVY)

    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(1)
    run_styled(p2, students, 'Calibri', 9, color=BLACK)

    p3 = cell.add_paragraph()
    p3.paragraph_format.space_after = Pt(2)
    run_styled(p3, org, 'Calibri', 8, italic=True, color=GREY)

    p4 = cell.add_paragraph()
    r = p4.add_run(f'  {result.upper()}  ')
    r.font.name = 'Calibri'
    r.font.size = Pt(7)
    r.font.bold = True
    r.font.color.rgb = WHITE
    rPr = r._r.get_or_add_rPr()
    rPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:fill="{accent}"/>'))

    add_spacer(4)

def swoc_box(title, items, bg_hex, accent_color):
    t = doc.add_table(rows=1, cols=1)
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, bg_hex)
    set_cell_margins(cell, top=100, bottom=100, left=160, right=120)
    set_cell_borders(cell,
        left={'color': accent_color, 'sz': '14', 'style': 'single'},
        top={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})

    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    run_styled(p, f'♦ {title}', 'Cambria', 11.5, bold=True, color=RGBColor.from_string(accent_color))

    for item in items:
        pi = cell.add_paragraph()
        pi.paragraph_format.space_before = Pt(1)
        pi.paragraph_format.space_after = Pt(1)
        run_styled(pi, f'  ›  {item}', 'Calibri', 9)

    add_spacer(4)

print("Compiling Master CE Annual Magazine 2025-26 DOCX...")

# ═════════════════════════════════════════════════════════════
#  COVER PAGE
# ═════════════════════════════════════════════════════════════
print("  → Building Cover Page")

cover = doc.add_table(rows=1, cols=1)
cover.alignment = WD_TABLE_ALIGNMENT.CENTER
make_table_full_width(cover)
cc = cover.cell(0, 0)
set_cell_shading(cc, NAVY_DARK_HEX)
set_cell_margins(cc, top=550, bottom=550, left=300, right=300)
set_cell_borders(cc)

# Logo
logo_path = os.path.join(IMG, 'audit_slide1_Google_Shape_119_p1.jpg')
if os.path.exists(logo_path):
    pl = cc.paragraphs[0]
    pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rl = pl.add_run()
    rl.add_picture(logo_path, width=Inches(0.95))

for _ in range(2):
    cc.add_paragraph().paragraph_format.space_after = Pt(2)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, "Bharatiya Vidya Bhavan’s", 'Cambria', 13, italic=True, color=GOLD)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'SARDAR PATEL INSTITUTE OF TECHNOLOGY', 'Cambria', 17, bold=True, color=WHITE)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
run_styled(p, '(Empowered Autonomous Institute Affiliated to University of Mumbai)', 'Calibri', 8.5, color=LIGHT_GREY)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, '─' * 42, 'Calibri', 8, color=GOLD)

for _ in range(2):
    cc.add_paragraph().paragraph_format.space_after = Pt(4)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'Department of', 'Cambria', 14, color=LIGHT_GREY)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'COMPUTER ENGINEERING', 'Cambria', 28, bold=True, color=WHITE)

for _ in range(3):
    cc.add_paragraph().paragraph_format.space_after = Pt(4)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, '─' * 42, 'Calibri', 8, color=GOLD)

cc.add_paragraph().paragraph_format.space_after = Pt(4)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'ANNUAL DEPARTMENTAL MAGAZINE', 'Calibri', 11, bold=True, color=GOLD)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, '2025 – 26', 'Cambria', 32, bold=True, color=GOLD)

for _ in range(3):
    cc.add_paragraph().paragraph_format.space_after = Pt(4)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'Bhavan’s Campus, Munshi Nagar, Andheri (West), Mumbai — 400 058', 'Calibri', 8.5, color=LIGHT_GREY)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'NBA ACCREDITED (Up to June 2027)  ·  NAAC "A" GRADE (Awarded June 2025)', 'Calibri', 8.5, bold=True, color=GOLD)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  TABLE OF CONTENTS (Hierarchical, Indented Index)
# ═════════════════════════════════════════════════════════════
print("  → Building Hierarchical Table of Contents")
add_spacer(14)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'TABLE OF CONTENTS', 'Cambria', 22, bold=True, color=NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'Department of Computer Engineering  ·  Academic Year 2025 — 2026', 'Calibri', 9, color=LIGHT_GREY)

gold_rule()

# (Number/Symbol, Title, is_subheading)
toc_structure = [
    ("01", "From the Principal’s Desk", False),
    ("02", "From the Head of Department’s Desk", False),
    ("03", "Vision and Mission Statement of the Department", False),
    ("04", "Program Educational Objectives (PEOs), Program Outcomes (POs) & PSOs", False),
    ("05", "About the Department — Curriculum and Syllabus Highlights", False),
    ("06", "Laboratory Facilities, Smart Classrooms & Major Equipment Added (2025–26)", False),
    ("07", "Overview of Academic Programs, Placements, Internships & Higher Studies", False),
    ("7.1", "Department Programs: Undergraduate, Post-Graduate & Doctoral", True),
    ("7.2", "Student Statistical Information & Cohort Analysis", True),
    ("7.3", "Campus Placements, Salary Tiers & Elite Recruitment Statistics", True),
    ("7.4", "Higher Studies Progression & Global University Admissions", True),
    ("08", "Faculty and Staff Information (Roster & Technical Staff)", False),
    ("09", "Additional Responsibilities to the Faculty", False),
    ("10", "Advisory Board, Board of Studies (BoS) & Department Academic Program Committee", False),
    ("11", "Consultancy, Grants, Patents & Research Publications", False),
    ("11.1", "Sponsored Research Proposals & Grant Submissions (₹6.05 Crore Portfolio)", True),
    ("11.2", "Industrial Consultancy Projects & Revenue Generation", True),
    ("11.3", "Faculty Publications, Citations (1,077+), Scopus & SCI/WoS Indexed Journals", True),
    ("11.4", "Student Research Publications & Intellectual Property (IPR)", True),
    ("12", "Faculty and Students Achievements", False),
    ("12.1", "Faculty Awards, Honors & NPTEL Discipline Stars", True),
    ("12.2", "Student Competition Participation: Hackathons, Case Studies & CTF", True),
    ("12.3", "National & International Victories: GATE DA AIR-97, Barclays, EY, Google Cloud", True),
    ("13", "List of Undergraduate Projects, M.Tech and Ph.D. Dissertations", False),
    ("13.1", "Undergraduate Capstone Projects & Emerging Domain Distribution", True),
    ("13.2", "High-Impact National Research Internships (NTRO Reverse Engineering)", True),
    ("13.3", "Post-Graduate (M.Tech) & Doctoral (Ph.D.) Research Tracks", True),
    ("14", "Industry–Institute Interaction", False),
    ("14.1", "Faculty Continuous Learning & Development (8,600+ Hours)", True),
    ("14.2", "Industry MoUs, Corporate Linkages & Technical Visits", True),
    ("14.3", "Workshops, STTPs, FDPs & Expert Talks Delivered by Faculty", True),
    ("15", "Distinguished Alumni of the Department & Entrepreneurs", False),
    ("15.1", "Alumni Participation in Internship Evaluations & Career Talks", True),
    ("15.2", "Industry Co-Teaching Collaborations (IIT Patna & Industry Leaders)", True),
    ("15.3", "Alumni Startup Founders, Innovators & Corporate Mentors", True),
    ("16", "Departmental Activities by Faculty, Students and Associations", False),
    ("16.1", "Parent–Teacher Meeting (PTM)", True),
    ("16.2", "FACE — Forum for Aspiring Computer Engineers", True),
    ("16.3", "CSI Student Chapter: Chronological Workshops & Hackathons", True),
    ("16.4", "IEEE Computer Society SPIT: Brain Drain, Research Ideation & AIRAVAT 3.0", True),
    ("17", "Student Support & Innovative Teaching-Learning Practices", False),
    ("17.1", "Structured Mentorship & 360-Degree Feedback Mechanism", True),
    ("17.2", "Targeted Support for Slow and Fast Learners", True),
    ("17.3", "Blooms Taxonomy Assessment, Gaussian Distribution & SA Analysis", True),
    ("17.4", "National Research Internships (IITs, BARC, ISRO, NTRO)", True),
    ("18", "Departmental Strengths, Weaknesses, Opportunities and Challenges (SWOC) & Goals", False),
    ("18.1", "Comprehensive SWOC Analysis", True),
    ("18.2", "Six Strategic Goals of the Department", True),
    ("18.3", "Key Departmental Strengths & Institutional Milestones", True),
    ("19", "Editorial Board & Meet the Team", False),
]

t_toc = doc.add_table(rows=len(toc_structure), cols=2)
t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
make_table_full_width(t_toc)
clear_all_borders(t_toc)

for i, (num, title, is_sub) in enumerate(toc_structure):
    c0, c1 = t_toc.cell(i, 0), t_toc.cell(i, 1)
    c0.width = Cm(1.6)

    p0 = c0.paragraphs[0]
    p1 = c1.paragraphs[0]

    if not is_sub:
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_styled(p0, num, 'Cambria', 10.5, bold=True, color=GOLD_DARK)
        run_styled(p1, title, 'Cambria', 10.5, bold=True, color=NAVY)
        set_cell_margins(c0, top=40, bottom=25, left=60, right=80)
        set_cell_margins(c1, top=40, bottom=25, left=60, right=60)
        set_cell_borders(c0, bottom={'color': BORDER_HEX, 'sz': '3', 'style': 'single'})
        set_cell_borders(c1, bottom={'color': BORDER_HEX, 'sz': '3', 'style': 'single'})
    else:
        p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_styled(p0, num, 'Calibri', 8.5, bold=False, color=LIGHT_GREY)
        # Indented subheading with arrow symbol
        p1.paragraph_format.first_line_indent = Cm(0.5)
        run_styled(p1, f'›   {title}', 'Calibri', 9.5, italic=False, color=BLACK)
        set_cell_margins(c0, top=15, bottom=15, left=60, right=80)
        set_cell_margins(c1, top=15, bottom=15, left=60, right=60)
        set_cell_borders(c0, bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'dotted'})
        set_cell_borders(c1, bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'dotted'})

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 1 — FROM THE PRINCIPAL'S DESK
# ═════════════════════════════════════════════════════════════
print("  → Building Section 1: Principal's Desk")
section_header("From the Principal’s Desk", "SECTION 01")

# Two-column layout for Principal: Photo on Left, Introduction & Lead Quote on Right
t_pr = doc.add_table(rows=1, cols=2)
t_pr.alignment = WD_TABLE_ALIGNMENT.CENTER
make_table_full_width(t_pr)
clear_all_borders(t_pr)

cp_photo = t_pr.cell(0, 0)
cp_photo.width = Cm(5.2)
set_cell_margins(cp_photo, top=40, bottom=40, left=40, right=60)
pp = cp_photo.paragraphs[0]
pp.alignment = WD_ALIGN_PARAGRAPH.CENTER

princ_img = os.path.join(IMG, 'principal_photo.png')
if os.path.exists(princ_img):
    rp = pp.add_run()
    rp.add_picture(princ_img, width=Inches(1.85))
    caption_p = cp_photo.add_paragraph()
    caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_p.paragraph_format.space_before = Pt(3)
    run_styled(caption_p, "Dr. Bhalchandra Chaudhari\nPrincipal, S.P.I.T.", 'Cambria', 9, bold=True, color=NAVY)

cp_lead = t_pr.cell(0, 1)
set_cell_margins(cp_lead, top=40, bottom=40, left=60, right=40)
pl1 = cp_lead.paragraphs[0]
pl1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
run_styled(pl1, "“To make aspiring minds confident and future-proof, education must prepare them for a marathon rather than a sprint. S.P.I.T. makes continuous, sincere efforts towards this.”", 'Cambria', 11, italic=True, color=NAVY)

pl2 = cp_lead.add_paragraph()
pl2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
pl2.paragraph_format.space_before = Pt(6)
run_styled(pl2, "I have great pleasure in expressing my thoughts as the Principal of Sardar Patel Institute of Technology (SPIT), the Numero Uno, self-financed, autonomous Institution of Maharashtra. We are a constituent of Bharatiya Vidya Bhavan, not just a conglomeration of more than 300 institutions, but a culture, a saga, a holy journey, started by Dr. K.M. Munshi with the support of Mahatma Gandhi in 1938. Imparting value-based education with Indian cultural ethos has always been the motto of Bhavan.", 'Calibri', 10)

add_spacer(6)

body("Engineers & technologists form the backbone of any nation’s economic development. The world is presently undergoing a very unprecedented, extraordinary, challenging time. A new normal is anticipated in many walks of life, including education. Such disruptions will come again and again in one or the other form. Engineers will work for 40–50 years of their life, they will have 3 to 4 diversified careers in technologies we are even unaware of. To make aspiring minds confident and future-proof, education must prepare them for a “marathon” rather than a “sprint”. SPIT makes continuous, sincere efforts towards this.", indent=True)

body("We focus on “How to learn?” rather than “What to learn?”. We believe in multidisciplinary exposure to the learners, yet ensuring growth in one vertical, cherishing human sensitivity and empathy. We have thoughtfully articulated a unique academic model towards this. Our splendid academic performance, sparkling placements (quantitative and qualitative), enrollment for higher studies at the best places of the world, prizes won by our students in national/international level technical competitions, in the past many years are the true testimonials for this.", indent=True)

body("With the support of a dedicated and hardworking faculty and staff, the institute has achieved enviable visibility and ranking in a short span. On behalf of all stakeholders of SPIT, I welcome you to this family and look forward to your valuable association with us for a better tomorrow. Four years of engineering education at Sardar Patel Institute of Technology or two years post-graduation, will undoubtedly empower you to lead a successful life.", indent=True)

pull_quote("Let’s grow together…", "Dr. Bhalchandra Chaudhari")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'Dr. Bhalchandra Chaudhari', 'Cambria', 12, bold=True, color=NAVY)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'Principal, Bharatiya Vidya Bhavan’s Sardar Patel Institute of Technology', 'Calibri', 9, color=GREY)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 2 — FROM THE HOD'S DESK
# ═════════════════════════════════════════════════════════════
print("  → Building Section 2: HOD's Desk")
section_header("From the Head of Department’s Desk", "SECTION 02")

stat_cards([
    ("98.21%", "Placement Rate (2025–26)", "green"),
    ("₹33 LPA", "Highest Package", "gold"),
    ("₹13.10 LPA", "Average CTC", "navy"),
    ("₹6.05 Cr", "Research Proposals", "gold"),
])

placeholder("INSERT HOD'S FORMAL PORTRAIT PHOTOGRAPH — Dr. Sudhir N. Dhage")
add_spacer(6)

body("It is with immense pride and deep satisfaction that I welcome you to the 2025–26 Annual Magazine of the Department of Computer Engineering at Bharatiya Vidya Bhavan's Sardar Patel Institute of Technology. This publication reflects the relentless commitment, intellectual rigor, and transformative achievements of our faculty, students, staff, and alumni community.", indent=True)

body_rich([
    ("The Academic Year 2025–26 has established monumental benchmarks across all criteria of departmental performance. Our graduating class achieved an outstanding placement rate of ", False),
    ("98.21%", True),
    (" with a highest compensation package of ", False),
    ("₹33.00 LPA", True),
    (" and an impressive average CTC of ", False),
    ("₹13.10 LPA", True),
    (" (median of ₹13.50 LPA). Our recruiters represent the apex of technology, finance, and product engineering — including ", False),
    ("Flipkart, Morgan Stanley, Goldman Sachs, Amazon, Deutsche Bank, VMware, MSCI, PhonePe, and JPMorgan Chase & Co.", True),
])

body_rich([
    ("Simultaneously, our culture of higher education and advanced research has achieved record heights. A remarkable ", False),
    ("36 graduating students (23% of the total cohort)", True),
    (" have advanced to premier global and national institutions including ", False),
    ("IIT Bombay, IIT Madras, IIT Patna, SPJIMR, and top-tier universities worldwide", True),
    (", nearly tripling our previous year's progression figure of 13 students.", False),
])

body_rich([
    ("Our faculty and students have driven extraordinary research output, producing ", False),
    ("43 research publications", True),
    (" (20 peer-reviewed international journal papers, 25 conference papers, 5 book chapters, and 2 published patents), achieving a ", False),
    ("+48% year-on-year expansion", True),
    (" with SCI/Web of Science journals surging by ", False),
    ("4x (from 1 to 4)", True),
    (" and Scopus-indexed journals tripling by ", False),
    ("3x (from 3 to 9)", True),
    (". Cumulative faculty citations currently stand at an exceptional ", False),
    ("1,077 citations", True),
    (". In parallel, our faculty members formulated and submitted ", False),
    ("10 high-value research proposals totaling ₹6.05 Crore", True),
    (" across prestigious funding schemes including Aspire LBI, DST-FIST, IEEE Emerging Grants, and the National Supercomputing Mission (NSM).", False),
])

body("Student innovation has been nothing short of exemplary. From Atharva Jagtap securing All India Rank 97 (AIR-97) in GATE Data Science & AI (GATE DA 2026), to national championship honors at the EY Techathon 6.0 (₹3,00,000/- prize), Barclays Hack-O-Hire, Google Cloud GenAI Exchange, NMIMS Adappt 4.0, and BITS Pilani — our students continue to prove that S.P.I.T. computer engineers stand at the national vanguard.", indent=True)

pull_quote("We prepare our graduates not merely to adapt to the future of technology, but to architect it with social responsibility, ethical grounding, and visionary innovation.", "Dr. Sudhir N. Dhage")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'Dr. Sudhir N. Dhage', 'Cambria', 12, bold=True, color=NAVY)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'Professor & Head, Department of Computer Engineering\nBharatiya Vidya Bhavan’s Sardar Patel Institute of Technology, Mumbai', 'Calibri', 9, color=GREY)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 3 & 4 — VISION, MISSION, PEOs, POs & PSOs
# ═════════════════════════════════════════════════════════════
print("  → Building Section 3 & 4: Strategic Foundations")
section_header("Vision, Mission & Program Outcomes", "SECTIONS 03 & 04")

sub_section_header("Strategic Foundation of the Department", "3.0")

vision_box("Department Vision",
    "To impart the education that graduates globally competent, socially sensitive, contributory computer professionals.")

doc.add_heading("Department Mission", level=2)
missions = [
    ("M1 — Pedagogical Excellence: ", "To provide a rigorous, conducive academic environment and constantly innovate in the design of the curriculum and its deployment through excellence in Teaching–Learning–Evaluation–Assessment processes."),
    ("M2 — Research & Industry Linkages: ", "To collaborate with industry and research organizations for knowledge creation and dissemination, fostering a culture of applied research and innovation among faculty and students."),
    ("M3 — Ethics & Holistic Leadership: ", "To inculcate professional ethics, social responsibility, and entrepreneurial thinking, preparing students to contribute meaningfully to society and the nation.")
]
for pfx, m_txt in missions:
    body_rich([(pfx, True, False, GOLD_DARK), (m_txt, False)])

gold_rule()

sub_section_header("Institutional Vision & Mission", "3.1")
highlight_card("Institute Vision",
    "To be a leading technical institution constantly pursuing excellence to develop holistic contributors for the society.",
    NAVY_HEX, LIGHT_BG)

bullet("To provide a conducive academic and research environment that stimulates learning, promotes intellectual growth, and ensures professional development of students, faculty, and staff.", "Academic Ecosystem: ")
bullet("To collaborate with renowned industry, academic institutions, and research laboratories for academics and research to solve real-life problems.", "Collaborative Research: ")
bullet("To leverage the strength of alumni to foster industry-relevant coaching and mentoring of the students.", "Alumni Mentorship: ")
bullet("To nurture values with a sense of social responsibility and inculcate attitudes and skills required for holistic development of the students.", "Value Education: ")

add_spacer(6)
sub_section_header("B.Tech. Program Educational Objectives (PEOs) & Program Specific Outcomes (PSOs)", "4.1")

peos = [
    ("PEO 1 — Technical Proficiency & Innovation: ", "Graduates will demonstrate proficiency in computing technologies and innovate in design, development, deployment, and maintenance of software systems serving various sectors of the economy."),
    ("PEO 2 — Leadership & Communication: ", "Graduates will be effective communicators, inspiring leaders, aspiring team members, and trailblazers of computational technologies."),
    ("PEO 3 — Lifelong Learning & Social Sensitivity: ", "Graduates will be contributory professionals demonstrating a lifelong learning attitude and deep sensitivity to societal challenges.")
]
for pfx, txt in peos:
    body_rich([(pfx, True, False, NAVY), (txt, False)])

add_spacer(4)
psos = [
    ("PSO 1 — Computational Core & Problem Solving: ", "Possess a strong foundation in Computer Engineering (Algorithms, Data Structures, AIML, Databases, Networks, Security, etc.) and apply the same to engineer feasible, scalable, and optimized solutions for real-world engineering challenges."),
    ("PSO 2 — Emerging Technologies & Professional Leadership: ", "Be motivated to engage in lifelong learning, adopt strong ethical values, and develop competencies in project and finance management, leadership, teamwork, and social awareness while demonstrating mastery across modern cloud platforms, cybersecurity architectures, and intelligent data systems.")
]
for pfx, txt in psos:
    body_rich([(pfx, True, False, GOLD_DARK), (txt, False)])

add_spacer(6)
sub_section_header("M.Tech. Program Outcomes (POs) & Program Educational Objectives (PEOs)", "4.2")

mtech_pos = [
    ("M.Tech PO 1: ", "An ability to independently carry out research / investigation and development work to solve practical engineering and scientific challenges."),
    ("M.Tech PO 2: ", "An ability to write and present a substantial, rigorous technical report and research documentation."),
    ("M.Tech PO 3: ", "Demonstrate an advanced degree of technical mastery over specialized computer engineering domains at a level significantly exceeding bachelor's degree standards.")
]
for pfx, txt in mtech_pos:
    body_rich([(pfx, True, False, NAVY), (txt, False)])

add_spacer(6)
sub_section_header("Accreditation Standing & Quality Certifications", "4.3")
data_table(
    ["Accrediting Body", "Accreditation Cycles / Duration", "Validity Status", "Significance"],
    [
        ["NBA Accreditation", "2013–14 (2 Years) · 2016–17 (3 Years) · 2021–22 (3 Years)", "Accredited up to June 30, 2027", "Tier-1 Autonomous Engineering Recognition"],
        ["NAAC Accreditation", "Cycle Assessment awarded June 28, 2025", "Grade 'A' (Valid until June 2027)", "Highest standard in institutional quality & governance"],
        ["UGC / Autonomous Status", "Conferred Empowered Autonomous Status", "Autonomous since 2017", "Full academic freedom for cutting-edge curriculum design"],
    ],
    [4.0, 5.0, 4.0, 4.0]
)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 5 — ABOUT THE DEPARTMENT & CURRICULUM
# ═════════════════════════════════════════════════════════════
print("  → Building Section 5: About the Department")
section_header("About the Department — Curriculum Highlights", "SECTION 05")

body("Established in 1995 with an initial sanctioned intake of 60 undergraduate students, the Department of Computer Engineering at S.P.I.T. has evolved into an academic center of national repute. Under autonomous governance, the department expanded its undergraduate capacity to 120 seats in 2021 and 240 seats in 2023, while sustaining vibrant post-graduate (M.Tech, Estd. 2011) and doctoral research centers (Ph.D., Estd. 2012, intake expanded to 30 seats in 2019).", indent=True)

data_table(
    ["Academic Program", "Branch / Specialization", "Year Established", "Current Intake", "Accreditation / Status"],
    [
        ["B.Tech.", "Computer Engineering", "1995 (60) · 2021 (120) · 2023", "240", "NBA Accredited up to June 2027"],
        ["M.Tech.", "Computer Engineering", "2011", "18", "Approved by AICTE / Affiliated to MU"],
        ["Ph.D.", "Computer Engineering", "2012 (10) · 2016 (20) · 2019", "30", "University of Mumbai Recognized Research Center"],
    ],
    [3.0, 4.5, 3.5, 2.0, 4.0]
)

sub_section_header("NEP 2020 Aligned Curriculum Structure & Credit Distribution", "5.1")
body("The department's curriculum is structured on the National Education Policy (NEP 2020) and the innovative Liberal 'Pi-Model', ensuring students gain deep vertical technical mastery alongside broad horizontal multidisciplinary exposure. The curriculum balances rigorous mathematical foundations, advanced algorithmic problem solving, hands-on laboratory experimentation, and industry immersion.", indent=True)

data_table(
    ["Curriculum Category", "Abbrev.", "Engagement Hours", "Percentage", "Educational Purpose"],
    [
        ["Lecture Engagement", "L", "102 Hours", "24.9%", "Fundamental theoretical principles & domain concepts"],
        ["Tutorial Engagement", "T", "3 Hours", "0.7%", "Small-group interactive numerical and analytical problem solving"],
        ["Laboratory Sessions", "L/P", "76 Hours", "18.5%", "Practical software engineering, implementation & benchmarking"],
        ["Self-Study / Project Learning", "O", "231 Hours", "56.3%", "Independent research, capstones, open-source & domain inquiry"],
        ["Total Cohort Engagement", "Total", "410 Hours", "100.0%", "Total Degree Credits Assigned: 160 Credits"],
    ],
    [4.5, 2.0, 3.2, 2.5, 4.8]
)

sub_section_header("Novel Pedagogical Pathways & Experiential Learning", "5.2")
bullet("Allows students to select specialized technical electives in Artificial Intelligence, Cloud Architectures, Quantum Computing, Blockchain, and Cybersecurity.", "Program Electives (PE): ")
bullet("Enables students to pursue minors in multidisciplinary fields including Management, Data Analytics, Electronics, and Humanities.", "Multidisciplinary Minors (MDM): ")
bullet("Undergraduate students participate in formal semester-long research internships at premier organizations such as NTRO, IIT Bombay, BARC, and ISRO.", "Research Internship Track: ")
bullet("Final-year students undergo six months of full-time industrial internship at premier technology enterprises with dual academic-industry evaluation.", "Industry Internship (6 Months): ")
bullet("Allows high-achieving learners to undertake specialized research dissertations culminating in international journal publications and patents.", "Honors with Research: ")
bullet("Mandatory integration of advanced NPTEL and SWAYAM MOOCs into elective baskets, fostering lifelong learning habits.", "MOOCs / NPTEL Integration: ")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 6 — LABORATORY FACILITIES & EQUIPMENT
# ═════════════════════════════════════════════════════════════
print("  → Building Section 6: Laboratories & Equipment")
section_header("Laboratory Facilities & Major Equipment (2025–26)", "SECTION 06")

body("The Department of Computer Engineering houses eight state-of-the-art computational laboratories, an enterprise-grade departmental server room, and high-performance GPU AI/ML compute infrastructure. In the academic year 2025–26, the department executed a major infrastructure upgrade, investing ₹49,44,388/- to deploy 80 brand-new high-performance workstations across four primary computing laboratories.", indent=True)

stat_cards([
    ("8", "Computing Laboratories", "navy"),
    ("80", "New Workstations Added", "gold"),
    ("₹49.44 L", "New Lab Investment", "green"),
    ("100%", "Gigabit LAN & WiFi", "navy"),
])

sub_section_header("Major Laboratory Equipment Added during 2025–26", "6.1")
data_table(
    ["Sr.", "Equipment Configuration & Specifications", "Qty", "Expenditure (INR)", "Target Laboratory & Room No."],
    [
        ["1", "HP ProDesk 2 GLA Desktop (R5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor)", "20", "Rs. 11,56,400/-", "Lab 5 (Room No. 606)"],
        ["2", "HP ProDesk 2 GLA Desktop (R5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor)", "20", "Rs. 11,56,400/-", "Lab 6 (Room No. 606)"],
        ["3", "HP ProDesk 2 GLA Desktop (Core i5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor)", "20", "Rs. 13,15,794/-", "Lab 4 (Room No. 604)"],
        ["4", "HP ProDesk 2 GLA Desktop (Core i5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor)", "20", "Rs. 13,15,794/-", "Lab 8 (Room No. 608)"],
        ["*", "Cumulative Investment in New Laboratory Hardware (2025–26)", "80", "Rs. 49,44,388/-", "Four Advanced Computing Labs"],
    ],
    [1.0, 7.5, 1.5, 3.5, 3.5]
)

sub_section_header("Laboratory Infrastructure Overview", "6.2")
labs_info = [
    ["Lab 1 & 2 (Room 601/602)", "Data Structures, Algorithms, Programming Methodology, Object-Oriented Technologies", "High-performance i5/i7 systems, Linux/Windows Dual Boot, IDEs, Git"],
    ["Lab 3 (Room 603)", "Database Management Systems, Big Data Analytics, Cloud Infrastructure", "Enterprise Oracle, PostgreSQL, MongoDB, Hadoop/Spark Clusters"],
    ["Lab 4 (Room 604)", "Operating Systems, System Programming, Computer Architecture", "HP ProDesk Core i5 Workstations, QEMU, GDB, Kernel Toolchains"],
    ["Lab 5 & 6 (Room 606)", "Artificial Intelligence, Deep Learning, Natural Language Processing, Computer Vision", "HP ProDesk R5 Systems, NVIDIA GPU Compute, PyTorch, TensorFlow"],
    ["Lab 7 (Room 607)", "Computer Networks, Cybersecurity, Cryptography & Network Defense", "Cisco Packet Tracer, Wireshark, Snort, Kali Linux virtualization"],
    ["Lab 8 (Room 608)", "Software Engineering, Cloud Computing, Capstone Projects", "HP ProDesk Core i5 Systems, Docker, Kubernetes, AWS/Azure SDKs"],
]
data_table(["Laboratory & Location", "Curricular & Research Domain", "Installed Hardware & Software Platforms"], labs_info, [3.8, 6.2, 7.0])

# Embed lab image if available
lab_img = os.path.join(IMG, 'doc_6._Lab_facilities,_Smart_Classroom_img4.png')
if os.path.exists(lab_img):
    add_image(lab_img, width=Inches(5.0), caption_text="Figure 6.1: Upgraded High-Performance Computing Laboratory at S.P.I.T. Department of Computer Engineering")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 7 — OVERVIEW OF ACADEMIC PROGRAMS & OUTCOMES
# ═════════════════════════════════════════════════════════════
print("  → Building Section 7: Placements & Higher Studies")
section_header("Academic Programs, Placements & Outcomes", "SECTION 07")

sub_section_header("Campus Placement Performance & Salary Tiers (2025–26)", "7.1")
body("The Department of Computer Engineering continued its unbroken record of premier recruitment during AY 2025–26, registering an outstanding 98.21% placement rate among eligible candidates. Out of 155 total students in the graduating cohort, 119 opted for campus placements and 110 secured lucrative offers across elite global corporations, investment banks, and technology leaders.", indent=True)

data_table(
    ["Academic Year", "Total Cohort", "Opted for Placement", "Students Placed", "Placement %", "Elite (≥18 LPA)", "Super Dream (9–17 LPA)", "Dream (6–9 LPA)", "Normal (<6 LPA)"],
    [
        ["2025–26", "155", "119", "110", "98.21%", "21", "70", "10", "9"],
        ["2024–25", "153", "130", "122", "97.60%", "47", "55", "14", "6"],
        ["2023–24", "78", "70", "66", "100.00%", "33", "30", "2", "1"],
    ],
    [2.0, 1.8, 2.2, 2.0, 1.8, 1.8, 2.5, 2.0, 1.9],
    compact=True
)

sub_section_header("Placement Salary Trends (2020–2026)", "7.2")
data_table(
    ["Compensation Metric (LPA)", "2020–21", "2021–22", "2022–23", "2023–24", "2024–25", "2025–26"],
    [
        ["Highest Salary Package", "42.00 LPA", "30.50 LPA", "110.00 LPA", "50.00 LPA", "50.00 LPA", "33.00 LPA"],
        ["Average Salary Package", "13.95 LPA", "14.77 LPA", "19.04 LPA", "17.33 LPA", "15.63 LPA", "13.10 LPA"],
        ["Median Salary Package", "13.00 LPA", "14.00 LPA", "14.00 LPA", "14.68 LPA", "14.00 LPA", "13.50 LPA"],
        ["Minimum Salary Package", "4.00 LPA", "6.30 LPA", "5.00 LPA", "3.00 LPA", "3.50 LPA", "3.00 LPA"],
    ],
    [4.2, 2.1, 2.1, 2.1, 2.1, 2.1, 2.1]
)

# Embed placement chart and recruiter logos
add_two_images(
    os.path.join(IMG, 'audit_updated_slide61_img8.jpeg'),
    "Figure 7.1: Placement Distribution by Salary Tiers (Elite, Super Dream, Dream, Normal)",
    os.path.join(IMG, 'audit_updated_slide24_img1.jpeg'),
    "Figure 7.2: Placement & Higher Studies Summary with Marquee Corporate Recruiters",
    width=Inches(2.7)
)

sub_section_header("Higher Studies Progression & Global University Admissions", "7.3")
body_rich([
    ("The department observed a tremendous surge in students opting for higher studies. During AY 2025–26, ", False),
    ("36 graduating students (23.2% of the entire cohort)", True),
    (" secured admissions to prestigious postgraduate and doctoral programs, up from ", False),
    ("13 students (8.5%) in 2024–25", True),
    (". This 177% increase was catalyzed by structured GRE/GATE coaching sessions, university application workshops, and alumni mentorship.", False)
])

data_table(
    ["Academic Year", "Graduating Cohort", "Higher Studies Progression", "Percentage of Class", "Target Academic Institutions"],
    [
        ["2025–26", "155 Students", "36 Students", "23.2%", "IIT Bombay, IIT Madras, IIT Patna, SPJIMR, Top US/European Universities"],
        ["2024–25", "153 Students", "13 Students", "8.5%", "Premier National & International Universities, DRDO, IIMs"],
        ["2023–24", "78 Students", "8 Students", "10.3%", "Reputed Technical & Management Universities"],
    ],
    [2.5, 3.0, 3.5, 2.5, 5.5]
)

# Embed overall outcome chart from slide 73
chart_outcomes = os.path.join(IMG, 'audit_updated_slide73_img1.jpeg')
if os.path.exists(chart_outcomes):
    add_image(chart_outcomes, width=Inches(5.0), caption_text="Figure 7.3: Department Overall Growth Snapshot & 2025–26 Graduate Outcomes (71% Placed, 23% Higher Studies)")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 8 — FACULTY AND STAFF INFORMATION
# ═════════════════════════════════════════════════════════════
print("  → Building Section 8: Faculty & Staff Roster")
section_header("Faculty and Staff Information", "SECTION 08")

body("The Department of Computer Engineering is powered by a highly accomplished team of 30 teaching faculty members, supported by skilled technical assistants and administrative staff. The faculty represents a rich confluence of academic pedagogy, doctoral research expertise, and industry experience.", indent=True)

# Faculty Roster Table from Slide 3
faculty_data = [
    ["1", "Dr. Sudhir N. Dhage", "Professor & Head of Department", "16", "Dr. Anuj Tawari", "Assistant Professor"],
    ["2", "Dr. Surekha Dholay", "Associate Professor", "17", "Ms. Soni Bhambar", "Assistant Professor"],
    ["3", "Dr. Anant V. Nimkar", "Associate Professor", "18", "Prof. Govind Gaundalkar", "Assistant Professor"],
    ["4", "Dr. Prasenjit Bhavathankar", "Associate Professor", "19", "Ms. Yogita Patil", "Assistant Professor"],
    ["5", "Prof. Anand Godbole", "Associate Professor", "20", "Prof. Azman Qureshi", "Assistant Professor"],
    ["6", "Dr. Nataasha Raul", "Assistant Professor", "21", "Prof. Jotsna Bhagat", "Assistant Professor"],
    ["7", "Dr. Pramod Bide", "Assistant Professor", "22", "Dr. Sanjuktarani Jena", "Assistant Professor"],
    ["8", "Dr. Kailas Devadkar", "Assistant Professor", "23", "Mr. Sudeep Mali", "Assistant Professor"],
    ["9", "Dr. Disha Sail", "Assistant Professor", "24", "Prof. Isha Sawalkar", "Assistant Professor"],
    ["10", "Dr. Jignesh Sisodia", "Assistant Professor", "25", "Mr. Ameya Nile", "Assistant Professor"],
    ["11", "Prof. Reeta Koshy", "Assistant Professor", "26", "Mr. Amey Morye", "Assistant Professor"],
    ["12", "Prof. Jyoti Ramteke", "Assistant Professor", "27", "Prof. Vishwas Badhe", "Assistant Professor"],
    ["13", "Prof. Sunil Ghane", "Assistant Professor", "28", "Ms. Aishwarya Nalawade", "Assistant Professor"],
    ["14", "Prof. Swapnali Kurhade", "Assistant Professor", "29", "Ms. Falak Malik", "Assistant Professor"],
    ["15", "Prof. Abhijeet Salunke", "Assistant Professor", "30", "Mr. Anas Ansari", "Assistant Professor"],
]

data_table(
    ["SN", "Faculty Member", "Designation", "SN", "Faculty Member", "Designation"],
    faculty_data,
    [0.8, 3.8, 3.0, 0.8, 3.8, 3.0],
    compact=True
)

sub_section_header("Technical Assistance & Supporting Staff", "8.1")
staff_data = [
    ["1", "Mr. Rajendra Nalawade", "Technical Assistant", "1", "Mr. Mangesh Birje", "Supporting Staff"],
    ["2", "Mr. Nandankumar Kinare", "Technical Assistant", "2", "Mr. Vishwanath Kamblekar", "Supporting Staff"],
    ["3", "Mrs. Rupali Bhosale", "Technical Assistant", "3", "Mr. Prathamesh Bhonde", "Supporting Staff"],
]
data_table(
    ["SN", "Technical Staff", "Designation", "SN", "Supporting Staff", "Designation"],
    staff_data,
    [0.8, 3.8, 3.0, 0.8, 3.8, 3.0],
    compact=True
)

# Department Group Photograph
fac_pic = os.path.join(BASE, 'Faclty group pic.JPG')
if os.path.exists(fac_pic):
    add_image(fac_pic, width=Inches(5.4), caption_text="Figure 8.1: Faculty and Staff Members of the Department of Computer Engineering, S.P.I.T. (AY 2025–26)")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 9 & 10 — FACULTY RESPONSIBILITIES & GOVERNANCE
# ═════════════════════════════════════════════════════════════
print("  → Building Section 9 & 10: Governance & Committees")
section_header("Governance, Portfolios & Advisory Bodies", "SECTIONS 09 & 10")

sub_section_header("Additional Institutional & Departmental Responsibilities", "9.0")
body("Faculty members actively contribute to institutional administration, accreditation cells, examination committees, and student development bodies, ensuring transparent governance and continual quality enhancement.", indent=True)

data_table(
    ["Key Institutional Portfolio", "Faculty In-Charge / Coordinator", "Key Responsibilities & Deliverables"],
    [
        ["Department NBA / NAAC Coordinator", "Dr. Surekha Dholay / Dr. Sudhir Dhage", "Curriculum mapping, criterion reporting, outcome assessment audits"],
        ["Department Academic Program Comm. (DAPC)", "Dr. Sudhir Dhage & Senior Faculty", "Curriculum structure, syllabus revisions, course outcome moderation"],
        ["Internal Quality Assurance Cell (IQAC)", "Dr. Prasenjit Bhavathankar / Dr. Pramod Bide", "Faculty appraisal, innovative teaching evaluation, quality audits"],
        ["Training & Placement Coordinators", "Dr. Kailas Devadkar / Prof. Sunil Ghane", "Corporate liaising, recruitment coordination, aptitude training"],
        ["Examination & Evaluation In-Charge", "Dr. Jignesh Sisodia / Prof. Govind Gaundalkar", "Blooms taxonomy audits, relative grading, moderation"],
        ["FACE & CSI Student Chapter Mentors", "Prof. Swapnali Kurhade / Prof. Anand Godbole", "Technical workshops, hackathons, coding contests, social drives"],
        ["IEEE Computer Society SPIT Mentors", "Dr. Sudhir Dhage / Dr. Pramod Bide", "Quantum conclave, AI hackathons, research ideation programs"],
    ],
    [4.2, 4.5, 8.3]
)

sub_section_header("Advisory Board, Board of Studies (BoS) & Academic Governance", "10.0")
bullet("The Board of Studies comprises senior faculty, distinguished university professors, industry leaders, and alumni representatives. It reviews syllabus content annually to incorporate cutting-edge developments in AI, Cloud, Quantum, and Security.", "Board of Studies (BoS): ")
bullet("The DAPC oversees semester-wise course attainment, assesses student feedback, and ensures continuous alignment with NBA/NAAC outcome indicators.", "Department Academic Program Committee: ")
bullet("The External Academic Audit conducts independent evaluations of question papers, laboratory rubrics, research publications, and curriculum compliance.", "External Academic Audit: ")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 11 — RESEARCH, GRANTS, PATENTS & PUBLICATIONS
# ═════════════════════════════════════════════════════════════
print("  → Building Section 11: Research & Grants")
section_header("Consultancy, Grants, Patents & Publications", "SECTION 11")

stat_cards([
    ("₹6.05 Cr", "Research Proposals (10)", "gold"),
    ("43", "Research Papers (2025–26)", "navy"),
    ("1,077", "Faculty Citations", "green"),
    ("₹10.2 L", "Consultancy Revenue", "gold"),
])

sub_section_header("Sponsored Research Proposals & Grant Submissions (₹6.05 Crore Portfolio)", "11.1")
body("During the 2024–25 and 2025–26 academic cycles, the Department of Computer Engineering demonstrated formidable momentum in external grant procurement, submitting 10 high-impact proposals totaling ₹6.05 Crore across major government and international bodies.", indent=True)

data_table(
    ["Sr.", "Funding Proposal Title & Sponsoring Agency", "Academic Year", "Grant Amount", "Faculty Principal Investigators / Focus"],
    [
        ["1", "Aspire LBI Grant Proposal", "2024–25", "Rs. 1.75 Crore", "Technology Business Incubation & Innovation Lab"],
        ["2", "DST-FIST Infrastructure Proposal", "2024–25", "Rs. 1.50 Crore", "Advanced Computing & Deep Tech Research Facilities"],
        ["3", "DST-FIST Infrastructure Proposal", "2025–26", "Rs. 1.50 Crore", "High-Performance Computing & GPU Acceleration Lab"],
        ["4", "IEEE Emerging Project Grant", "2024–25", "Rs. 48.0 Lakh", "Applied Intelligent Systems & Edge Computing"],
        ["5", "IEEE Emerging Project Grant", "2025–26", "Rs. 48.0 Lakh", "AI/ML Solutions for Socio-Economic Impact"],
        ["6", "ISEC 2027 Conference Hosting Grant", "2025–26", "Rs. 15.0 Lakh", "International Software Engineering Conference Hosting"],
        ["7", "NSM Hackathon (Sponsored by C-DAC & WCE)", "2025–26", "Rs. 6.0 Lakh", "National Supercomputing Mission Parallel Computing Contest"],
        ["8", "National Supercomputing Mission (NSM) FDP Proposal", "2025–26", "Rs. 5.0 Lakh", "Faculty Capacity Building in HPC & Quantum Algorithms"],
        ["9", "MeitY & EICT Faculty Development Programme", "2025–26", "Rs. 5.0 Lakh", "Next-Gen AI & Embedded Systems STTP"],
        ["10", "IEEE Quantum Computing Conclave 2026", "2025–26", "Rs. 3.0 Lakh", "National Quantum Technology Symposium & Industry Meet"],
        ["*", "Cumulative Value of Submitted Funding Proposals", "2024–26", "Rs. 6.05 Crore", "10 High-Value Institutional Research Proposals"],
    ],
    [0.8, 6.0, 1.8, 2.5, 6.0]
)

# Embed funding chart from slide 42
add_image(
    os.path.join(IMG, 'audit_updated_slide42_img4.jpeg'),
    width=Inches(5.2),
    caption_text="Figure 11.1: Comprehensive Proposal Submission & Funding Portfolio (2024–25 to 2025–26) Representing ₹6.05 Crore across 10 Proposals"
)

sub_section_header("Industrial Projects & Consultancy Revenue", "11.2")
body_rich([
    ("The department actively converts technical capabilities into industrial solutions. In AY 2024–25, ", False),
    ("Dr. Anant V. Nimkar", True),
    (" successfully spearheaded the industrial consultancy project titled ", False),
    ("“IoT GATEWAY DESIGN”", True),
    (", generating commercial revenue of ", False),
    ("Rs. 10.2 Lakhs", True),
    (" for the institute. Additional industrial collaborations and technical advisory roles continue to expand across automotive, fintech, and embedded systems sectors.", False)
])

sub_section_header("Faculty Publications, Citations & Quality Journals", "11.3")
body("Research output expanded by +48% in one year, marked by a decisive transition into premier high-impact factor journals. SCI/Web of Science publications quadrupled from 1 to 4, and Scopus-indexed journal papers tripled from 3 to 9.", indent=True)

data_table(
    ["Academic Year", "Journal Papers", "Conference Papers", "Book Chapters", "Patents", "Citations", "Scopus Journals", "SCI / WoS"],
    [
        ["2025–26", "20", "25", "5", "2", "1,077", "09", "04"],
        ["2024–25", "13", "25", "2", "5", "1,168", "03", "01"],
        ["YoY Growth", "+53.8%", "Sustained", "+150.0%", "Active", "Cumulative", "+200.0% (3x)", "+300.0% (4x)"],
    ],
    [2.5, 2.0, 2.2, 2.0, 1.5, 1.8, 2.3, 1.8]
)

# Embed research chart from slide 71
add_image(
    os.path.join(IMG, 'audit_updated_slide71_img1.jpeg'),
    width=Inches(5.0),
    caption_text="Figure 11.2: Research & Industry Interaction Strong Upward Trend (43 Total Research Output, 4x in SCI/WoS, 3x in Scopus)"
)

sub_section_header("Student Research Publications & Intellectual Property (IPR)", "11.4")
data_table(
    ["Student Category", "2024–25 Journals", "2024–25 Conferences", "2025–26 Journals", "2025–26 Conferences", "Research Focus & Domains"],
    [
        ["Undergraduate (UG)", "9", "18", "10", "1", "Deep Learning, NLP, Edge Computing, Blockchain"],
        ["Post-Graduate (M.Tech)", "1", "2", "2", "29", "Advanced Vision, Parallel Systems, Cloud Architectures"],
        ["Doctoral (Ph.D.)", "2", "2", "2", "0", "Quantum Computing, Medical Informatics, Security"],
    ],
    [3.5, 2.5, 2.5, 2.5, 2.5, 3.5]
)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 12 — FACULTY AND STUDENTS ACHIEVEMENTS
# ═════════════════════════════════════════════════════════════
print("  → Building Section 12: Achievements & Victories")
section_header("Faculty & Students Achievements", "SECTION 12")

sub_section_header("Faculty Awards, Honors & National Recognitions", "12.1")
data_table(
    ["Honored Faculty Member", "Prestigious Award / Recognition Title", "Conferring Organization / Context"],
    [
        ["Prof. Anand Godbole", "NPTEL Discipline Star Award", "National Programme on Technology Enhanced Learning (NPTEL)"],
        ["Dr. Pramod Bide", "Best Innovative Teaching Award", "L&T EduTech Excellence in Engineering Education"],
        ["Dr. Jignesh Sisodia", "Best Software Project Award", "Cantilever Industry Project Showcase"],
        ["Prof. Govind Gaundalkar", "Best Question Paper Award (Digital Systems)", "S.P.I.T. Academic Examination Audit Committee"],
        ["Dr. Jignesh Sisodia & Prof. Abhijit Salunke", "Best Question Paper Award (CCN)", "Computer Communication and Networks Course Audit"],
    ],
    [4.5, 5.0, 7.5]
)

sub_section_header("Student Participation in Competitive Arenas", "12.2")
body("During AY 2025–26, Computer Engineering students demonstrated massive engagement across technical competitions, registering 50 entries across 20 distinct hackathons, 23 entries across 13 case study competitions, and 87 participants in co-curricular technical symposiums.", indent=True)

data_table(
    ["Competition Arena / Category", "Distinct Events", "Total Student Entries", "Representative Victories / Outcomes"],
    [
        ["National & Corporate Hackathons", "20 Events", "50 Entries", "EY Techathon National Champions, Barclays Hack-O-Hire, Adappt 4.0"],
        ["Case Study & Strategy Competitions", "13 Events", "23 Entries", "Corporate consulting challenges, IIM Case Contests"],
        ["Technical & National Examination", "5 Events", "9 Entries", "GATE DA All India Rank 97 (Score 816)"],
        ["Innovation & Project Recognition", "7 Events", "9 Entries", "Google Cloud GenAI Exchange 3rd Runner Up"],
        ["Security / Bug Bounty / CTF", "4 Events", "4 Entries", "Ethereum Bug Bounty & Zoho Hall of Fame entries"],
    ],
    [4.5, 2.5, 2.5, 7.5]
)

sub_section_header("Marquee Student Victories & National Honors (Slide 60 Showcase)", "12.3")

# Embedded Student Achievement Photos (Slide 60)
add_two_images(
    os.path.join(IMG, 'audit_updated_slide60_img4.jpeg'),
    "Figure 12.1: Atharva Sharad Jagtap — All India Rank 97 (AIR-97) in GATE DA 2026 (Score: 816, IIT Guwahati)",
    os.path.join(IMG, 'audit_updated_slide60_img6.jpeg'),
    "Figure 12.2: Winners of Barclays Hack-O-Hire Hackathon at Gera Commerzone, Pune",
    width=Inches(2.7)
)

add_two_images(
    os.path.join(IMG, 'audit_updated_slide60_img7.jpeg'),
    "Figure 12.3: Google Cloud GenAI Exchange Hackathon — 3rd Runner Up (Team GG_Hustlers: Career & Skills Advisor)",
    os.path.join(IMG, 'audit_updated_slide60_img9.jpeg'),
    "Figure 12.4: National Champions — EY Techathon 6.0 (Ernst & Young People's Choice Award — ₹3,00,000/- Prize)",
    width=Inches(2.7)
)

add_two_images(
    os.path.join(IMG, 'audit_updated_slide60_img10.jpeg'),
    "Figure 12.5: Adappt 4.0 Hackathon Winners (1st Place Trophy, NMIMS)",
    os.path.join(IMG, 'audit_updated_slide60_img8.jpeg'),
    "Figure 12.6: First Position Winners at National Hackathon Hosted by BITS Pilani, Rajasthan",
    width=Inches(2.7)
)

achievement_card(
    "EY Techathon 6.0 — National Champion & People’s Choice Award (₹3,00,000/-)",
    "Team Dwayne Fernandes, S.P.I.T. Computer Engineering",
    "Ernst & Young Global Consulting Hackathon",
    "1st Place National Champions",
    GOLD_HEX
)

achievement_card(
    "GATE 2026 — Data Science & Artificial Intelligence (AIR-97)",
    "Atharva Sharad Jagtap (Score: 816 / 1000 | Marks: 65.0 / 100)",
    "IIT Guwahati / National GATE Committee",
    "All India Rank 97",
    NAVY_HEX
)

achievement_card(
    "Barclays Hack-O-Hire Hackathon Winner",
    "S.P.I.T. Computer Engineering Student Team",
    "Barclays Global Service Centre, Pune",
    "Winner Trophy & Pre-Placement Offers",
    GREEN_HEX
)

achievement_card(
    "Google Cloud GenAI Exchange Hackathon",
    "Team GG_Hustlers (Personalized Career and Skills Advisor)",
    "Google Cloud India",
    "3rd Runner Up Nationally",
    NAVY_HEX
)

achievement_card(
    "Adappt 4.0 Hackathon Winner & BITS Pilani 1st Position",
    "S.P.I.T. Student Teams",
    "NMIMS Mumbai & BITS Pilani Rajasthan",
    "1st Place Winners",
    GOLD_HEX
)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 13 — UNDERGRADUATE PROJECTS & RESEARCH
# ═════════════════════════════════════════════════════════════
print("  → Building Section 13: Capstone Projects & Domains")
section_header("Undergraduate Projects & Research Tracks", "SECTION 13")

sub_section_header("Capstone Project Domain Evolution (2024–25 vs 2025–26)", "13.1")
body("Analysis of undergraduate project allocations highlights a definitive strategic shift toward high-complexity Artificial Intelligence and Deep Learning systems. In 2025–26, Deep Learning and AI accounted for 31 projects (over 54% of all capstones), supported by Machine Learning (12), Blockchain (4), Image Processing (3), Web Applications (5), and Network Security (2).", indent=True)

# Side-by-side charts from slide 20
add_two_images(
    os.path.join(IMG, 'audit_updated_slide20_img4.jpeg'),
    "Figure 13.1: Domain-wise Project Distribution (AY 2024–25)",
    os.path.join(IMG, 'audit_updated_slide20_img5.jpeg'),
    "Figure 13.2: Domain-wise Project Distribution (AY 2025–26: Massive Surge in Deep Learning & AI)",
    width=Inches(2.7)
)

data_table(
    ["Project Technology Domain", "2024–25 Count", "2025–26 Count", "Growth Trend", "Strategic Realignment"],
    [
        ["Deep Learning", "12 Projects", "16 Projects", "+33.3%", "Computer vision, generative models, transformer systems"],
        ["Artificial Intelligence", "8 Projects", "15 Projects", "+87.5%", "Autonomous agents, neural architectures, LLM systems"],
        ["Machine Learning", "13 Projects", "12 Projects", "Consolidated", "Predictive modeling, healthcare & financial analytics"],
        ["Web & Distributed Applications", "13 Projects", "5 Projects", "Transitioned", "Shift from basic CRUD to intelligent AI-driven platforms"],
        ["Blockchain & Web3", "0 Projects", "4 Projects", "New Growth", "Smart contracts, decentralized security, verifiable credentials"],
        ["Image Processing & Vision", "2 Projects", "3 Projects", "+50.0%", "Biometric processing, medical diagnostics"],
        ["Network Security & Cryptography", "3 Projects", "2 Projects", "Active", "Threat modeling, zero-trust protocols"],
    ],
    [4.0, 2.5, 2.5, 2.5, 5.5]
)

sub_section_header("High-Impact Defense & National Research Internships", "13.2")
body_rich([
    ("The department places special emphasis on solving real-world national problems. In AY 2025–26, students ", False),
    ("Raj Kalpesh Mathuria (UID: 2023300139)", True),
    (" and ", False),
    ("Vishwajit Sambhaji Sarnobat (UID: 2023300195)", True),
    (" completed an advanced research project titled ", False),
    ("“Proprietary File System Reverse Engineering and Data Extraction Framework”", True),
    (" at the ", False),
    ("National Technical Research Organisation (NTRO)", True),
    (" under the expert mentorship of ", False),
    ("Sameera Mishra", True),
    (". Such projects demonstrate the critical defense and intelligence relevance of the department's training.", False)
])

data_table(
    ["UID", "Student Name", "Project Title", "Hosting Organization", "External Mentor"],
    [
        ["2023300139", "Raj Kalpesh Mathuria", "Proprietary File System Reverse Engineering and Data Extraction Framework", "NTRO", "Sameera Mishra"],
        ["2023300195", "Vishwajit Sambhaji Sarnobat", "Proprietary File System Reverse Engineering and Data Extraction Framework", "NTRO", "Sameera Mishra"],
    ],
    [2.5, 4.0, 5.5, 2.5, 2.5]
)

sub_section_header("Post-Graduate & Doctoral Research Themes", "13.3")
bullet("Focuses on large-scale parallel processing, distributed computing, explainable AI (XAI), and zero-trust cloud network security.", "M.Tech. Research: ")
bullet("Advanced investigations in quantum machine learning, medical image synthesis, post-quantum cryptography, and intelligent edge frameworks.", "Ph.D. Dissertations: ")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 14 — INDUSTRY-INSTITUTE INTERACTION
# ═════════════════════════════════════════════════════════════
print("  → Building Section 14: Industry Interaction")
section_header("Industry–Institute Interaction & Faculty Growth", "SECTION 14")

sub_section_header("Faculty Development: Sharp Growth Across the Board (8,600+ Hours)", "14.1")
body("In AY 2025–26, Computer Engineering faculty members dedicated over 8,600 hours to continuous pedagogical, technological, and research upskilling across 166 training programs received, 71 training programs imparted, 61 outreach engagements as resource persons, and 19 MOOC certifications.", indent=True)

data_table(
    ["Faculty Development Metric", "2024–25 Count", "2025–26 Count", "Growth (%)", "2025–26 Total Hours", "2025–26 Total Days"],
    [
        ["Training Programs Received (FDP/STTP)", "55", "166", "+201.8%", "4,621 Hours", "1,235 Days"],
        ["Training Imparted to External Peers", "20", "71", "+255.0%", "1,788 Hours", "155 Days"],
        ["Outreach as Keynote / Resource Person", "15", "61", "+306.7%", "1,442 Hours", "90 Days"],
        ["NPTEL / Coursera MOOCs Completed", "10", "19", "+90.0%", "825 Hours", "828 Days"],
        ["*Cumulative Faculty Growth Indicators", "100", "317", "+217.0%", "8,676 Hours", "2,308 Days"],
    ],
    [5.0, 2.2, 2.2, 2.2, 2.8, 2.6]
)

# Embed faculty growth chart from slide 72
add_image(
    os.path.join(IMG, 'audit_updated_slide72_img1.jpeg'),
    width=Inches(5.0),
    caption_text="Figure 14.1: Faculty Development Sharp Growth Across the Board (+202% Training Received, +255% Training Imparted, +307% Outreach)"
)

sub_section_header("Industry Collaborations, MoUs & Corporate Linkages", "14.2")
bullet("Formal collaboration on student training, hackathons, and curriculum alignment with national computing frameworks.", "C-DAC & Western Regional Centres: ")
bullet("Joint technology development in IoT and embedded communications generating commercial revenue.", "Industry Consultancy Linkages: ")
bullet("Expert lecture series and specialized curriculum delivery in cloud computing and enterprise cybersecurity.", "Corporate Technology Partners: ")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 15 — DISTINGUISHED ALUMNI & ENTREPRENEURS
# ═════════════════════════════════════════════════════════════
print("  → Building Section 15: Alumni & Entrepreneurs")
section_header("Distinguished Alumni & Entrepreneurs", "SECTION 15")

sub_section_header("Alumni Participation in Academic Evaluation & Career Mentorship", "15.1")
body_rich([
    ("The department maintains deep, continuous ties with its global alumni network. In AY 2024–25, ", False),
    ("18 alumni experts", True),
    (" and in AY 2025–26, ", False),
    ("11 alumni leaders", True),
    (" actively participated in the rigorous evaluation of final-year students undergoing industry and research internships, providing pragmatic corporate feedback.", False)
])

highlight_card(
    "Distinguished Alumni Speaker Session: Consulting & Management Careers",
    "Dr. Sudhir Dhage organized a premier alumni leadership session on “Exploring and Preparing for Management and Consulting Careers”, delivered by alumnus Soumyadeb Bhattacharjee, Strategy & Management Consultant. The session guided over 120 students on consulting case preparation, strategy frameworks, and executive decision-making.",
    GOLD_HEX, GOLD_BG_HEX
)

sub_section_header("Industry Co-Teaching Collaborations", "15.2")
highlight_card(
    "Academic Co-Teaching with Premier Institutions: IIT Patna",
    "Under the leadership of Prof. Swapnali Kurhade, the department instituted co-teaching for the advanced Artificial Intelligence & Machine Learning (AIML) course, featuring specialized guest instruction by Dr. Maheshkumar H. Kolekar from IIT Patna (conducted on 13/10/2024). This initiative exposed students to IIT-level algorithmic rigor and research methodologies.",
    NAVY_HEX, LIGHT_BG
)

sub_section_header("Alumni Pioneers & Startup Founders", "15.3")
body("Graduates of the Department of Computer Engineering have founded impactful startups and occupy leadership positions across Google, Microsoft, Meta, Amazon, Goldman Sachs, and top academic institutions worldwide. They regularly mentor student founders and sponsor student hackathons.", indent=True)

placeholder("INSERT DETAILED ALUMNI PROFILES & STARTUP PROFILES IF ADDITIONAL DETAILS AVAILABLE")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 16 — DEPARTMENTAL ACTIVITIES & ASSOCIATIONS
# ═════════════════════════════════════════════════════════════
print("  → Building Section 16: Departmental Activities")
section_header("Departmental Activities & Student Chapters", "SECTION 16")

sub_section_header("Parent–Teacher Meeting (PTM)", "16.1")
body("The department conducts structured Parent–Teacher Meetings to review student academic progress, attendance, and career preparedness with parents and guardians. Feedback received from parents is incorporated into departmental action plans.", indent=True)

ptm_pic = os.path.join(BASE, 'PTM.JPG')
if os.path.exists(ptm_pic):
    add_image(ptm_pic, width=Inches(5.0), caption_text="Figure 16.1: Parent–Teacher Meeting Conducted by Department of Computer Engineering, S.P.I.T.")

sub_section_header("FACE — Forum for Aspiring Computer Engineers", "16.2")
bullet("Organized social donation drives mobilizing books, stationery, and educational supplies for underprivileged students.", "Donation Drives 1 & 2: ")
bullet("High-intensity intra-departmental competitive coding and algorithmic tournament.", "FaceOff: ")
bullet("Portfolio and project showcase platform for students to demonstrate web, app, and AI projects.", "Codefolio: ")

sub_section_header("CSI Student Chapter: Chronological Workshops & Hackathons (2024–26)", "16.3")
body("The Computer Society of India (CSI) student chapter conducted an extensive suite of hands-on technical workshops, developer masterclasses, and national hackathons across both academic years:", indent=True)

data_table(
    ["Academic Year", "Event Name", "Execution Date", "Domain & Technical Focus"],
    [
        ["2025–26", "Git and GitHub Workshop", "1st October 2025", "Version control, collaborative workflows, open source"],
        ["2025–26", "Code Housie", "16th October 2025", "Interactive algorithmic problem-solving game"],
        ["2025–26", "SomETHing Decentralised", "3rd November 2025", "Web3, Solidity smart contracts, decentralized apps"],
        ["2025–26", "Pixel Perfect", "14th February 2026", "UI/UX design principles, responsive web layouts"],
        ["2025–26", "Code it Ralph", "16th February 2026", "Competitive speed coding and algorithmic optimization"],
        ["2025–26", "ML Sprint", "17th February 2026", "Fast-paced machine learning modeling and inference"],
        ["2025–26", "CTRL Fix + Takedown", "18th February 2026", "Debugging tournament, system repair, edge-case testing"],
        ["2025–26", "Design Disaster", "19th February 2026", "Critique and redesign of complex user interfaces"],
        ["2025–26", "UnSQLved", "20th February 2026", "Advanced database query optimization and SQL logic"],
        ["2025–26", "S.P.I.T. Hackathon 2026", "21st February 2026", "Flagship national 24-hr hackathon (3,300+ applicants)"],
        ["2025–26", "SE Hackathon", "18th April 2026", "Software engineering architecture and full-stack deploy"],
        ["2024–25", "Tech Roadmap & GSoC Session", "30th Sep – 1st Oct 2024", "Career navigation and Google Summer of Code guidance"],
        ["2024–25", "Source Start & Minithon", "Oct 2024", "Month-long open source sprint & rapid hackathon"],
        ["2024–25", "Frontend Odyssey & rAPId", "1st – 5th Feb 2025", "Modern web frontend frameworks and API engineering"],
        ["2024–25", "S.P.I.T. Hackathon 2025", "8th–9th February 2025", "National software innovation hackathon"],
    ],
    [1.8, 3.8, 3.2, 8.2],
    compact=True
)

sub_section_header("IEEE Computer Society SPIT: High-Impact Conclaves & Hackathons", "16.4")
body("IEEE CS-SPIT represents a cornerstone of the department's technical excellence, driving cutting-edge research conclaves and national hackathons:", indent=True)

# Embed IEEE CS photos from slide 58 and 59
add_two_images(
    os.path.join(IMG, 'audit_updated_slide58_img4.jpeg'),
    "Figure 16.2: IEEE CS-SPIT — 'Brain Drain' Icebreaker (Nov 2025) and 'Research Ideation' with Dr. Fitri Yakub (Jan 2026)",
    os.path.join(IMG, 'audit_updated_slide59_img4.jpeg'),
    "Figure 16.3: IEEE CS-SPIT — Quantum Computing Conclave (Feb 2026) and Flagship AIRAVAT 3.0 AI Hackathon (Apr 2026)",
    width=Inches(2.7)
)

bullet("Conducted on 17th November 2025 with a ₹4,000+ prize pool as an engaging icebreaker for incoming first-year students to introduce computational problem solving.", "Brain Drain (17 Nov 2025): ")
bullet("Hosted on January 28, 2026 in Room 008 with 80+ participants. Featured international virtual instruction by Dr. Fitri Yakub (IEEE Malaysia / MJIIT) alongside Dr. Sudhir Dhage and Dr. Pramod Bide on scientific problem formulation and research publishing.", "Research Ideation Workshop (28 Jan 2026): ")
bullet("Organized on 20–21 February 2026, uniting renowned scientists, quantum researchers, faculty, and students to examine quantum algorithms, post-quantum cryptography, and physical qubits.", "Quantum Computing Conclave (20–21 Feb 2026): ")
bullet("The third edition of IEEE CS-SPIT's flagship AI hackathon conducted on April 11–12, 2026. A 24-hour national-level competition assembling the country's brightest young innovators to engineer generative AI and autonomous systems.", "AIRAVAT 3.0 (11–12 April 2026): ")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 17 — STUDENT SUPPORT & INNOVATIVE TEACHING
# ═════════════════════════════════════════════════════════════
print("  → Building Section 17: Student Support & TLE")
section_header("Student Support & Innovative Teaching", "SECTION 17")

sub_section_header("Structured Mentorship Scheme & 360-Degree Feedback", "17.1")
body("The department operates a structured mentorship framework ensuring every student is paired with a dedicated faculty mentor. Regular one-on-one sessions monitor academic progress, attendance, mental well-being, and career planning. Concerns are formally logged in the Issue Register and addressed via documented Action Taken Reports. A 360-degree feedback system gathers inputs from students on curriculum, teaching, and infrastructure.", indent=True)

sub_section_header("Differentiated Learning: Support for Slow and Fast Learners", "17.2")
bullet("Identified through continuous diagnostic evaluations, mid-semester exams, and assignment rubrics. Supported via remedial lectures, peer mentoring, study notes on Moodle, and special assistance for differently-abled students.", "Slow Learners: ")
bullet("Encouraged to pursue honors courses, research mini-projects, competitive coding, hackathons, and research internships at premier defense and academic centers.", "Fast Learners: ")

sub_section_header("Innovative Teaching-Learning & Evaluation (TLE) Practices", "17.3")
bullet("All question papers are designed strictly according to Bloom's Taxonomy, ensuring systematic coverage of Course Outcomes (COs) and balanced distribution between lower and higher-order thinking skills.", "Bloom's Taxonomy Framework: ")
bullet("For End Semester Examinations (ESE), two distinct sets of question papers are formulated per course, accompanied by rigorous internal moderation and external academic audits.", "Two-Set ESE Question Papers: ")
bullet("MSE and ESE grading is calibrated to yield a natural Gaussian curve. Faculty convene formal departmental review meetings to evaluate Subject Average (SA) graphs (e.g. Operating Systems course with 451 enrolled students analyzed to establish an SA benchmark of 73, ensuring 10–15% students attain the AA grade).", "Gaussian Distribution & SA Calibration: ")
bullet("Following MSE and ESE result compilation, an Open House is conducted where students physically inspect their evaluated answer scripts, review rubrics, and seek immediate resolution of discrepancies.", "Open House Transparency: ")
bullet("The Internal Quality Assurance Cell (IQAC) conducts transparent annual evaluations to recognize faculty members who implement standout innovations in pedagogy and evaluation.", "IQAC Best Performing Teacher Recognition: ")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 18 — SWOC ANALYSIS & STRATEGIC ROADMAP
# ═════════════════════════════════════════════════════════════
print("  → Building Section 18: SWOC & Goals")
section_header("SWOC Analysis & Department Strategic Goals", "SECTION 18")

sub_section_header("Comprehensive SWOC Analysis (Slide 78)", "18.1")

swoc_box("STRENGTHS (Established Academic Ecosystem)", [
    "NEP 2020-aligned, outcome-based curriculum with flexible electives and experiential learning pathways.",
    "Active research, IPR filing, student capstones, and robust higher-studies progression (23% of graduating cohort).",
    "Deep industry and alumni engagement in curriculum co-teaching, internship evaluation, and placement mentorship.",
    "Consistently sparkling placement statistics (98.21% placement rate in 2025–26, ₹33 LPA highest, ₹13.10 LPA average CTC)."
], LIGHT_BG, NAVY_HEX)

swoc_box("AREAS TO STRENGTHEN / WEAKNESSES (Scale & Systematize)", [
    "Deepen sustained sponsored research, consultancy revenue, and external project funding from national agencies.",
    "Further strengthen systematic documentation of mentorship sessions and continuous-improvement tracking.",
    "Broaden international academic exchanges and collaborative research across overseas universities."
], GOLD_BG_HEX, GOLD_HEX)

swoc_box("OPPORTUNITIES (Emerging & Collaborative Domains)", [
    "Scale advanced learning and laboratory infrastructure in AI/ML, Generative AI, Cloud, Cybersecurity, and Quantum Computing.",
    "Expand research-institute collaborations with IITs, BARC, ISRO, and defense bodies like NTRO.",
    "Strengthen student entrepreneurship, hackathons, and startup incubation converting prototypes into registered patents."
], GREEN_BG, GREEN_HEX)

swoc_box("CHALLENGES (Rapidly Changing Ecosystem)", [
    "Keeping curriculum, laboratory infrastructure, and faculty competencies continuously aligned with hyper-fast technology cycles.",
    "Scaling personalized experiential learning, lab infrastructure, and mentorship amid increasing student intake.",
    "Compete effectively for high-value research grants, premier faculty talent, and exclusive corporate partnerships."
], RED_BG, RED_HEX)

add_spacer(6)
sub_section_header("Six Strategic Goals of the Department (Slide 79)", "18.2")
goals = [
    ("Goal 01 — Research & Sponsored Projects: ", "Expand funded research, industrial consultancy, IPR filings, and interdisciplinary project activity across government and global bodies."),
    ("Goal 02 — Industry–Academia Integration: ", "Deepen corporate co-teaching, sponsored project labs, semester-long internships, and executive expert engagement."),
    ("Goal 03 — Emerging Technology Readiness: ", "Continuously strengthen computational infrastructure for AI/ML, Generative AI, Quantum Computing, Cloud, and Cybersecurity."),
    ("Goal 04 — Innovation & Entrepreneurship: ", "Convert high-quality student projects and hackathons into viable prototypes, research publications, IPR, and incubation ventures."),
    ("Goal 05 — Global & Higher-Study Engagement: ", "Expand defense/research internships, academic partnerships, and structured pathways to premier global and national higher-study institutions."),
    ("Goal 06 — Continuous Quality Improvement: ", "Leverage academic audits, 360-degree feedback, and outcome evidence to continuously upgrade mentorship, labs, and pedagogical processes.")
]
for pfx, g_txt in goals:
    body_rich([(pfx, True, False, NAVY), (g_txt, False)])

add_spacer(6)
sub_section_header("Core Departmental Strengths & Milestones (Slide 80)", "18.3")
data_table(
    ["Pillar", "Strategic Dimension", "Demonstrated Institutional Milestone (AY 2025–26)"],
    [
        ["01", "Academic Excellence", "Autonomous, NEP 2020-aligned curriculum, Liberal Pi-Model, 160 credits with experiential tracks"],
        ["02", "Faculty Excellence", "30 experienced faculty members; over 8,600 hours invested in FDPs, MOOCs, and expert outreach"],
        ["03", "Research & Innovation", "43 research papers (20 journals, 25 conferences, 5 book chapters, 2 patents, 1,077 citations)"],
        ["04", "Industry & Alumni Connect", "Active alumni evaluation of internships, co-teaching with IIT Patna, ₹10.2L consultancy project"],
        ["05", "Student Outcomes", "98.21% placement rate, ₹33 LPA highest, 23% cohort advancing to IITs and world-class universities"],
        ["06", "Learning Infrastructure", "₹49.44 Lakh invested in 80 new workstations across Labs 4, 5, 6, and 8; GPU computing facilities"],
    ],
    [1.0, 4.0, 11.0]
)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 19 — MEET THE TEAM / EDITORIAL BOARD
# ═════════════════════════════════════════════════════════════
print("  → Building Section 19: Editorial Team")
section_header("Editorial Board & Meet the Team", "SECTION 19")

body("The 2025–26 Annual Magazine of the Department of Computer Engineering is the culmination of months of dedicated collaboration among faculty mentors, student editors, designers, photographers, and student contributors.", indent=True)

data_table(
    ["Editorial Board Portfolio", "Designation / Role", "Department Affiliation"],
    [
        ["Patron & Chief Mentor", "Dr. Bhalchandra Chaudhari", "Principal, S.P.I.T."],
        ["Head of Department & Editor-in-Chief", "Dr. Sudhir N. Dhage", "Professor & Head, Computer Engineering"],
        ["Faculty Editorial Advisors", "Prof. Swapnali Kurhade · Dr. Pramod Bide", "Assistant Professors, Computer Engineering"],
        ["Student Chief Editors", "Student Editorial Leads (CE 2025–26)", "Department Student Editorial Committee"],
        ["Visual Design & Creative Direction", "Creative & Design Team (CE 2025–26)", "Department Creative Arts & Media Cell"],
        ["Data Curation & Audit Synthesis", "Editorial Documentation Team", "External Audit Documentation Support"],
    ],
    [4.5, 5.5, 6.0]
)

add_spacer(8)
pull_quote(
    "We express our heartfelt gratitude to the faculty, staff, alumni, and students whose remarkable contributions bring these pages to life. Here's to another year of pushing boundaries and engineering the future!",
    "The CE Magazine Editorial Team"
)

placeholder("ADD STUDENT EDITORS & DESIGN TEAM NAMES/PHOTOS IF DESIRED")

# Save document
print(f"Saving final master document to {OUT}...")
doc.save(OUT)
print("SUCCESS: CE_Annual_Magazine_2025-26_v2.docx created successfully!")
