"""
CE Department Annual Magazine 2025-26 — PREMIUM DOCX Generator (v2)
Bharatiya Vidya Bhavan's Sardar Patel Institute of Technology, Mumbai

Generates a visually stunning, editable Word document matching the HTML design.
Uses table-based layout techniques for colored boxes, stat cards, and section headers.
"""

import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

BASE = r'c:\Users\Nirupam\Documents\magazine ce'
IMG = os.path.join(BASE, 'extracted_images')
OUT = os.path.join(BASE, 'CE_Annual_Magazine_2025-26_v2.docx')

# ── Colors ──
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
# PAGE SETUP
# ═══════════════════════════════════════
sec = doc.sections[0]
sec.page_width = Cm(21)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(2.0)
sec.right_margin = Cm(2.0)

# ── Default Style ──
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10.5)
style.font.color.rgb = BLACK
style.paragraph_format.space_after = Pt(4)
style.paragraph_format.line_spacing = 1.15

for level, (sz, sp_b) in {1: (20, 0), 2: (14, 14), 3: (12, 10)}.items():
    h = doc.styles[f'Heading {level}']
    h.font.name = 'Cambria'
    h.font.size = Pt(sz)
    h.font.bold = True
    h.font.color.rgb = NAVY
    h.paragraph_format.space_before = Pt(sp_b)
    h.paragraph_format.space_after = Pt(4)


# ═══════════════════════════════════════
# UTILITY FUNCTIONS
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


# ═══════════════════════════════════════
# DESIGN COMPONENTS
# ═══════════════════════════════════════

def section_header(title, subtitle=None):
    """Full-width navy header band with gold bottom accent — matches HTML design."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, NAVY_DARK_HEX)
    set_cell_margins(cell, top=140, bottom=120, left=200, right=200)
    set_cell_borders(cell,
        bottom={'color': GOLD_HEX, 'sz': '18', 'style': 'single'})
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if subtitle:
        ps = cell.paragraphs[0]
        ps.paragraph_format.space_after = Pt(1)
        run_styled(ps, subtitle, 'Calibri', 8, color=GOLD, bold=True)
        # Add the title paragraph
        pt = cell.add_paragraph()
        pt.paragraph_format.space_before = Pt(2)
        pt.paragraph_format.space_after = Pt(2)
        run_styled(pt, title, 'Cambria', 19, bold=True, color=WHITE)
    else:
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        run_styled(p, title, 'Cambria', 19, bold=True, color=WHITE)

    add_spacer(8)


def stat_cards(stats):
    """Row of stat boxes with large numbers — like HTML stat-cards."""
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
        run_styled(p, str(value), 'Cambria', 24, bold=True, color=color)

        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        run_styled(p2, label.upper(), 'Calibri', 7.5, bold=True, color=GREY)

    add_spacer(8)


def pull_quote(text):
    """Gold left-border quote — using a 2-column table with narrow gold stripe."""
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)

    # Gold stripe column
    gc = t.cell(0, 0)
    gc.width = Cm(0.3)
    set_cell_shading(gc, GOLD_HEX)
    set_cell_borders(gc)
    gc.paragraphs[0].text = ''

    # Quote text
    qc = t.cell(0, 1)
    set_cell_shading(qc, GOLD_BG_HEX)
    set_cell_margins(qc, top=100, bottom=100, left=160, right=160)
    set_cell_borders(qc)
    p = qc.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    run_styled(p, f'\u201c{text}\u201d', 'Cambria', 11, italic=True, color=NAVY)

    # Set column widths
    tbl = t._tbl
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if len(cols) >= 2:
            cols[0].set(qn('w:w'), '200')
            cols[1].set(qn('w:w'), '9600')

    add_spacer(6)


def highlight_box(content_fn, accent_color=GOLD_HEX, bg_color=GOLD_BG_HEX):
    """Colored box with left accent border."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, bg_color)
    set_cell_margins(cell, top=100, bottom=100, left=180, right=160)
    set_cell_borders(cell,
        left={'color': accent_color, 'sz': '18', 'style': 'single'},
        top={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})
    content_fn(cell)
    add_spacer(6)


def vision_box(label, text):
    """Navy background box for vision/mission — like HTML vision-box."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_shading(cell, NAVY_DARK_HEX)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    set_cell_borders(cell,
        bottom={'color': GOLD_HEX, 'sz': '12', 'style': 'single'})

    pl = cell.paragraphs[0]
    pl.paragraph_format.space_after = Pt(4)
    run_styled(pl, label.upper(), 'Calibri', 8, bold=True, color=GOLD)

    pt = cell.add_paragraph()
    pt.paragraph_format.space_before = Pt(2)
    run_styled(pt, f'\u201c{text}\u201d', 'Cambria', 13, italic=True, color=WHITE)

    add_spacer(8)


def data_table(headers, rows, col_widths=None, compact=False):
    """Professional data table with navy header and alternating rows."""
    sz = 8.5 if compact else 9.5
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    t.style = 'Table Grid'

    # Header
    for j, h in enumerate(headers):
        c = t.cell(0, j)
        set_cell_shading(c, NAVY_HEX)
        set_cell_margins(c, top=50, bottom=50, left=80, right=80)
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        run_styled(p, h, 'Calibri', sz - 0.5, bold=True, color=WHITE)

    # Rows
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i + 1, j)
            if i % 2 == 1:
                set_cell_shading(c, ALT_ROW)
            set_cell_margins(c, top=30, bottom=30, left=80, right=80)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            # First column in data bold
            is_bold = (j == 0 and len(headers) > 2) or str(val).startswith('*')
            val_text = str(val).lstrip('*')
            run_styled(p, val_text, 'Calibri', sz, bold=is_bold)

    if col_widths:
        for j, w in enumerate(col_widths):
            for row_obj in t.rows:
                row_obj.cells[j].width = Cm(w)

    add_spacer(6)
    return t


def company_tags(companies, elite_set=None):
    """Company tags as a paragraph of styled inline elements."""
    elite_set = elite_set or set()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    for i, c in enumerate(companies):
        sep = '   ·   ' if i > 0 else ''
        if sep:
            run_styled(p, sep, 'Calibri', 8, color=LIGHT_GREY)
        is_elite = c in elite_set
        run_styled(p, c, 'Calibri', 9, bold=is_elite, color=NAVY if is_elite else GREY)
    add_spacer(4)


def gold_rule():
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_borders(cell,
        bottom={'color': GOLD_HEX, 'sz': '8', 'style': 'single'})
    cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    cell.paragraphs[0].paragraph_format.line_spacing = Pt(1)
    set_cell_margins(cell, 0, 0, 0, 0)
    add_spacer(6)


def thin_rule():
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    cell = t.cell(0, 0)
    set_cell_borders(cell,
        bottom={'color': BORDER_HEX, 'sz': '4', 'style': 'single'})
    cell.paragraphs[0].paragraph_format.space_after = Pt(0)
    cell.paragraphs[0].paragraph_format.line_spacing = Pt(1)
    set_cell_margins(cell, 0, 0, 0, 0)
    add_spacer(4)


def body(text, indent=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.8)
    run_styled(p, text, 'Calibri', 10.5, italic=italic)
    return p


def body_rich(segments, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    for seg in segments:
        txt, bold = seg[0], seg[1] if len(seg) > 1 else False
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
    r = p.add_run(f'  \u25B6 {text}  ')
    r.font.name = 'Calibri'
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0xCC, 0x33, 0x33)
    r.bold = True
    rPr = r._r.get_or_add_rPr()
    rPr.append(parse_xml(f'<w:highlight {nsdecls("w")} w:val="yellow"/>'))


def add_image(path, width=Inches(5.2)):
    full = path if os.path.isabs(path) else os.path.join(BASE, path)
    if os.path.exists(full):
        try:
            doc.add_picture(full, width=width)
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            return True
        except:
            placeholder(f'IMAGE ERROR: {os.path.basename(path)}')
            return False
    else:
        placeholder(f'IMAGE NOT FOUND: {os.path.basename(path)}')
        return False


def caption(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    run_styled(p, text, 'Calibri', 8, italic=True, color=GREY)


def achievement_card(event, students, org, result, accent=GOLD_HEX):
    """Styled achievement entry with colored left border."""
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
    # Result badge
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
    run_styled(p, f'\u2666 {title}', 'Cambria', 12, bold=True, color=RGBColor.from_string(accent_color))

    for item in items:
        pi = cell.add_paragraph()
        pi.paragraph_format.space_before = Pt(1)
        pi.paragraph_format.space_after = Pt(1)
        run_styled(pi, f'  \u203A  {item}', 'Calibri', 9)

    add_spacer(4)


print("Building premium magazine DOCX (v2)...")

# ═════════════════════════════════════════════════════════════
#  COVER PAGE
# ═════════════════════════════════════════════════════════════
print("  \u2192 Cover Page")

# Full-page navy cover using a single table cell
cover = doc.add_table(rows=1, cols=1)
cover.alignment = WD_TABLE_ALIGNMENT.CENTER
make_table_full_width(cover)
cc = cover.cell(0, 0)
set_cell_shading(cc, NAVY_DARK_HEX)
set_cell_margins(cc, top=600, bottom=600, left=300, right=300)
set_cell_borders(cc)

# Logo
logo_path = os.path.join(IMG, 'audit_slide1_Google_Shape_119_p1.jpg')
if os.path.exists(logo_path):
    pl = cc.paragraphs[0]
    pl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rl = pl.add_run()
    rl.add_picture(logo_path, width=Inches(0.9))

# Institution
for _ in range(2):
    cc.add_paragraph().paragraph_format.space_after = Pt(2)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, "Bharatiya Vidya Bhavan\u2019s", 'Cambria', 13, italic=True, color=GOLD)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'SARDAR PATEL INSTITUTE OF TECHNOLOGY', 'Cambria', 16, bold=True, color=WHITE)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(4)
run_styled(p, '(Empowered Autonomous Institute Affiliated to University of Mumbai)', 'Calibri', 8, color=LIGHT_GREY)

# Divider line
p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, '\u2500' * 40, 'Calibri', 8, color=GOLD)

for _ in range(2):
    cc.add_paragraph().paragraph_format.space_after = Pt(4)

# Department
p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'Department of', 'Cambria', 14, color=LIGHT_GREY)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'COMPUTER ENGINEERING', 'Cambria', 28, bold=True, color=WHITE)

for _ in range(3):
    cc.add_paragraph().paragraph_format.space_after = Pt(4)

# Divider
p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, '\u2500' * 40, 'Calibri', 8, color=GOLD)

cc.add_paragraph().paragraph_format.space_after = Pt(4)

# Annual Magazine
p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'ANNUAL MAGAZINE', 'Calibri', 11, bold=True, color=GOLD)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, '2025 \u2013 26', 'Cambria', 32, bold=True, color=GOLD)

for _ in range(3):
    cc.add_paragraph().paragraph_format.space_after = Pt(4)

# Bottom info
p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'Munshi Nagar, Andheri (West), Mumbai \u2014 400 058', 'Calibri', 8, color=LIGHT_GREY)

p = cc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'NBA Accredited \u2014 2013 \u00b7 2016 \u00b7 2021 \u00b7 2024', 'Calibri', 8, bold=True, color=GOLD)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  TABLE OF CONTENTS
# ═════════════════════════════════════════════════════════════
print("  \u2192 Table of Contents")
add_spacer(20)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'CONTENTS', 'Cambria', 24, bold=True, color=NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'Academic Year 2025 \u2014 2026', 'Calibri', 9, color=LIGHT_GREY)

gold_rule()

toc_items = [
    ("01", "From the Principal\u2019s Desk"),
    ("02", "From the Head of Department\u2019s Desk"),
    ("03", "Vision and Mission Statement"),
    ("04", "Program Educational Objectives & Program Outcomes"),
    ("05", "About the Department \u2014 Curriculum & Syllabus Highlights"),
    ("06", "Laboratory Facilities, Smart Classrooms & Equipment"),
    ("07", "Overview of Academic Programs \u2014 Placements, Internships & Higher Studies"),
    ("08", "Faculty and Staff"),
    ("09", "Additional Responsibilities to the Faculty"),
    ("10", "Advisory Board and Board of Studies (BoS), DAPC"),
    ("11", "Consultancy, Grants, Patents & Research Publications"),
    ("12", "Faculty and Students Achievements"),
    ("13", "List of Undergraduate Projects, M.Tech and Ph.D."),
    ("14", "Industry\u2013Institute Interaction"),
    ("15", "Distinguished Alumni of the Department"),
    ("16", "Departmental Activities \u2014 FACE, CSI, IEEE & Associations"),
    ("17", "Student Support & Innovative Teaching"),
    ("18", "Departmental SWOC Analysis"),
    ("\u2605", "Meet the Team"),
]

t = doc.add_table(rows=len(toc_items), cols=3)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
make_table_full_width(t)
clear_all_borders(t)

for i, (num, title) in enumerate(toc_items):
    c0, c1, c2 = t.cell(i, 0), t.cell(i, 1), t.cell(i, 2)
    c0.width = Cm(1.2)
    c2.width = Cm(1.0)

    # Number
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_styled(p0, num, 'Cambria', 10, bold=True, color=GOLD_DARK)

    # Title
    p1 = c1.paragraphs[0]
    is_last = i == len(toc_items) - 1
    run_styled(p1, title, 'Calibri', 10.5, italic=is_last)

    # Dotted bottom border on each row
    for c in [c0, c1, c2]:
        set_cell_borders(c, bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'dotted'})
        set_cell_margins(c, top=30, bottom=30, left=60, right=60)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 1 — PRINCIPAL'S MESSAGE
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 1: Principal\u2019s Message")
section_header("From the Principal\u2019s Desk", "SECTION 01")

placeholder("INSERT PRINCIPAL\u2019S PHOTOGRAPH HERE")
add_spacer(6)

body("It gives me immense pleasure to present the Annual Magazine of the Department of Computer Engineering for the academic year 2025\u201326. This publication is a reflection of the collective endeavour, intellectual vigour, and creative spirit that define our institution.", indent=True)
body("Sardar Patel Institute of Technology has always been committed to nurturing engineers who are not only technically proficient but also socially conscious and globally aware. The Department of Computer Engineering stands as a testament to this commitment, consistently setting benchmarks in academics, research, industry collaboration, and student development.", indent=True)
body("This year has been particularly noteworthy. The department has continued its legacy of excellence with a remarkable placement record, significant research contributions published in leading journals and conferences, and students who have distinguished themselves at national and international platforms \u2014 from ICPC to EY Techathon to the Ethereum Bug Bounty Hall of Fame.", indent=True)
body('Our curriculum, aligned with the National Education Policy 2020 through the innovative Liberal "Pi-Model," empowers students with the flexibility to explore diverse domains ranging from Artificial Intelligence and Blockchain to Quantum Computing and Cybersecurity, while remaining firmly grounded in engineering fundamentals.', indent=True)
body("I congratulate the Head of Department, the faculty members, the supporting staff, and above all, the students, for their relentless pursuit of excellence. I am confident that the Department of Computer Engineering will continue to scale greater heights in the years to come.", indent=True)

add_spacer(6)
placeholder("EDIT / REPLACE THE ABOVE MESSAGE WITH THE ACTUAL PRINCIPAL\u2019S MESSAGE IF AVAILABLE")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'The Principal', 'Cambria', 12, bold=True, color=NAVY)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'Sardar Patel Institute of Technology, Mumbai', 'Calibri', 9, color=GREY)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 2 — HOD'S MESSAGE
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 2: HOD\u2019s Message")
section_header("From the Head of Department\u2019s Desk", "SECTION 02")

placeholder("INSERT HOD\u2019S PHOTOGRAPH \u2014 Dr. Sudhir N. Dhage")
add_spacer(6)

body("It is with great pride and satisfaction that I present the Annual Magazine of the Department of Computer Engineering for the academic year 2025\u201326. This report encapsulates the achievements, progress, and aspirations of our department across academics, research, placements, and co-curricular engagement.", indent=True)

body_rich([
    ("The year 2025\u201326 has been a landmark year. Our placement percentage of ", False),
    ("98.21%", True), (", highest CTC of ", False),
    ("\u20b933 LPA", True), (", and average CTC of ", False),
    ("\u20b910.60 LPA", True),
    (" demonstrate the trust leading companies place in our graduates. Recruiters include Barclays, Nomura, BlackRock, JP Morgan, PhonePe, Morgan Stanley, and many more. ", False),
    ("32 students", True),
    (" have secured admissions to IIT Bombay, IIT Madras, IIT Patna, SPJIMR, and DRDO.", False),
])

body_rich([
    ("On the research front: ", False),
    ("15 international journal publications", True),
    (" and ", False), ("25 conference papers", True),
    (" (9 Scopus-indexed, 3 SCIE). We filed a patent for an AI-powered interactive voice-enabled drawing tool.", False),
])

body("Our students won the EY Techathon 6.0, qualified as ICPC Asia Amritapuri Regionalists, earned spots on the Ethereum and Zoho Bug Bounty Halls of Fame, and won at IIT Bombay, IIM Indore, and Google Cloud hackathons. The SPIT Hackathon attracted 3,316 registrations.", indent=True)

body("I extend my heartfelt gratitude to our faculty, staff, industry partners, alumni, parents, and students.", indent=True)

add_spacer(6)
placeholder("EDIT / REPLACE WITH ACTUAL HOD\u2019S MESSAGE IF AVAILABLE")

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'Dr. Sudhir N. Dhage', 'Cambria', 12, bold=True, color=NAVY)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
run_styled(p, 'Professor & Head, Dept. of Computer Engineering\nS.P.I.T., Mumbai', 'Calibri', 9, color=GREY)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 3 & 4 — VISION, MISSION, PEOs, POs
# ═════════════════════════════════════════════════════════════
print("  \u2192 Sections 3 & 4: Vision, Mission, PEOs & POs")
section_header("Vision, Mission & Program Outcomes", "SECTIONS 03 & 04")

vision_box("Our Vision",
    "To impart the education that graduates globally competent, socially sensitive, contributory computer professionals.")

doc.add_heading("Our Mission", level=2)
for i, m in enumerate([
    "To provide a rigorous, conducive academic environment and constantly innovate in the design of the curriculum and its deployment through excellence in Teaching\u2013Learning\u2013Evaluation\u2013Assessment processes.",
    "To collaborate with industry and research organisations for knowledge creation and dissemination, fostering a culture of applied research and innovation among faculty and students.",
    "To inculcate professional ethics, social responsibility, and entrepreneurial thinking, preparing students to contribute meaningfully to society and the nation.",
], 1):
    body_rich([(f"M{i}:  ", True, False, GOLD_DARK), (m, False)])

placeholder("ADD OR EDIT VISION/MISSION STATEMENTS IF DIFFERENT")
gold_rule()

doc.add_heading("Program Educational Objectives (PEOs)", level=2)
for pf, txt in [
    ("PEO 1: ", "Graduates will have a successful career in the IT industry, academia, research, or entrepreneurial ventures."),
    ("PEO 2: ", "Graduates will pursue lifelong learning and professional development, adapting to evolving technologies."),
    ("PEO 3: ", "Graduates will exhibit professionalism, ethical values, effective communication, and social responsibility."),
]:
    bullet(txt, bold_prefix=pf)

thin_rule()

doc.add_heading("Program Outcomes (POs)", level=2)
for pf, txt in [
    ("PO 1 \u2014 Engineering Knowledge: ", "Apply knowledge of mathematics, science, and engineering to solve complex problems."),
    ("PO 2 \u2014 Problem Analysis: ", "Identify, formulate, and analyse complex engineering problems using first principles."),
    ("PO 3 \u2014 Design & Development: ", "Design solutions meeting specified needs with societal consideration."),
    ("PO 4 \u2014 Investigation: ", "Use research-based knowledge to investigate complex problems and synthesise information."),
    ("PO 5 \u2014 Modern Tool Usage: ", "Apply appropriate techniques, resources, and modern engineering tools."),
    ("PO 6 \u2014 Engineer & Society: ", "Apply contextual reasoning to assess societal, health, safety, and cultural issues."),
    ("PO 7 \u2014 Sustainability: ", "Understand the societal and environmental impact of engineering solutions."),
    ("PO 8 \u2014 Ethics: ", "Apply ethical principles and commit to professional practice norms."),
    ("PO 9 \u2014 Teamwork: ", "Function effectively as individual, member, or leader in diverse teams."),
    ("PO 10 \u2014 Communication: ", "Communicate effectively, write reports, and make presentations."),
    ("PO 11 \u2014 Project Management: ", "Apply management principles in multidisciplinary environments."),
    ("PO 12 \u2014 Life-long Learning: ", "Engage in independent and life-long learning."),
]:
    bullet(txt, bold_prefix=pf)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 5 — ABOUT THE DEPARTMENT
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 5: About the Department")
section_header("About the Department", "SECTION 05 \u2014 Curriculum & Syllabus Highlights")

doc.add_heading("Redefining Engineering Education: The New-Age Curriculum", level=2)
p = doc.add_paragraph()
run_styled(p, "From Learning Technology to Creating Impact", 'Cambria', 11, italic=True, color=GREY)

body("Engineering education is undergoing a fundamental transformation. The engineer of today is expected to think critically, communicate effectively, work across disciplines, understand emerging technologies, and contribute responsibly to society and industry.")

body_rich([
    ("Recognising this, S.P.I.T. has introduced its Liberal ", False),
    ("\u201cPi-Model\u201d", True),
    (" of Engineering Education under NEP 2020. The curriculum nurtures well-rounded professionals with interdisciplinary awareness, ethical grounding, and adaptability.", False),
])

pull_quote("The 160-credit structure distributes learning across Basic & Engineering Sciences, Programme Core, Programme Electives, Experiential Learning, Open Electives, and Skill Enhancement Courses \u2014 providing 57 different pathways to graduation.")

doc.add_heading("From Core Computing to Emerging Technologies", level=3)
body("Initial semesters build foundations in Calculus, Linear Algebra, Probability, Digital Systems, Programming, and Data Structures. Core subjects follow: Discrete Structures, DBMS, Computer Organisation, OS, Software Engineering, DAA. The fifth semester introduces AI & ML, Cryptography, Theory of Computation, and Distributed Computing.")

doc.add_heading("Specialisation Pathways", level=3)

def spec_content(cell):
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    run_styled(p, '\U0001f916 AI & Data Science: ', 'Calibri', 10, bold=True, color=NAVY)
    run_styled(p, 'Deep Learning, Generative AI, Agentic AI, Explainable AI, NLP, Big Data Analytics, Quantum ML', 'Calibri', 9.5)
    p2 = cell.add_paragraph()
    p2.paragraph_format.space_after = Pt(4)
    run_styled(p2, '\U0001f512 Cybersecurity: ', 'Calibri', 10, bold=True, color=NAVY)
    run_styled(p2, 'Ethical Hacking, Blockchain, Digital Forensics, Information Security', 'Calibri', 9.5)
    p3 = cell.add_paragraph()
    p3.paragraph_format.space_after = Pt(4)
    run_styled(p3, '\u26a1 Systems & Hardware: ', 'Calibri', 10, bold=True, color=NAVY)
    run_styled(p3, 'Embedded Systems, IoT, Robotics, TinyML, HPC, Cloud Computing', 'Calibri', 9.5)
    p4 = cell.add_paragraph()
    run_styled(p4, '\U0001f3a8 Visual & Interactive: ', 'Calibri', 10, bold=True, color=NAVY)
    run_styled(p4, 'Computer Vision, AR/VR, Visual Computing, Quantum Computing', 'Calibri', 9.5)

highlight_box(spec_content, NAVY_HEX, BLUE_SOFT)

doc.add_heading("Multidisciplinary Minors & Beyond", level=3)
bullet("Banking Technology (Barclays), Management (SPJIMR), Financial & Strategic Mgmt", bold_prefix="Industry MDMs: ")
bullet("Industrial IoT, VLSI & Signal Processing, UX/Interface Design", bold_prefix="Cross-Domain: ")
bullet("Yoga, Bhagavad Gita, Photography, Film, Entrepreneurship", bold_prefix="Liberal Learning: ")

doc.add_heading("Experiential Learning", level=3)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(6)
run_styled(p, "Learn \u2192 Explore \u2192 Build \u2192 Experiment \u2192 Solve \u2192 Create", 'Cambria', 12, bold=True, color=NAVY)

body("Students undertake Mini Projects I & II (Sem V\u2013VI), a Main Project, and research or industry internships. An Honours by Research pathway (18 extra credits, 8.25+ CGPA) provides sustained research experience at the undergraduate level.")

pull_quote("The Liberal Pi-Model makes the academic journey more flexible and holistic \u2014 preparing professionals who can use technology thoughtfully to address the challenges of tomorrow.")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 6 — LAB FACILITIES
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 6: Lab Facilities")
section_header("Laboratory Facilities & Smart Classrooms", "SECTION 06")

body("The Department maintains state-of-the-art laboratory infrastructure across the fourth, fifth, and sixth floors, ensuring modern computing resources for all academic and project-related needs.")

doc.add_heading("Department Laboratories", level=3)
for img, cap in [
    ('ptm_slide38_Google_Shape_610_p104.jpg', '604 \u2014 Distributed Computing Laboratory'),
    ('ptm_slide38_Google_Shape_612_p104.jpg', '606 \u2014 Data Analytics Laboratory'),
    ('ptm_slide38_Google_Shape_614_p104.jpg', '507 \u2014 Operating Systems Laboratory'),
    ('ptm_slide38_Google_Shape_616_p104.jpg', '509 \u2014 Project Laboratory'),
]:
    add_image(os.path.join(IMG, img), width=Inches(4.0))
    caption(cap)

doc.add_heading("Smart Classrooms", level=3)
for img in ['ptm_slide39_Google_Shape_626_p105.jpg', 'ptm_slide39_Google_Shape_627_p105.jpg']:
    add_image(os.path.join(IMG, img), width=Inches(4.0))
    caption("ICT-Enabled Smart Classroom")

doc.add_heading("Major Equipment Added in 2025\u201326", level=3)
def equip_content(cell):
    p = cell.paragraphs[0]
    run_styled(p, '40 New HP Desktop Computers ', 'Calibri', 10.5, bold=True, color=NAVY)
    run_styled(p, 'installed in Database Lab (Lab-5) & Data Analytics Lab (Lab-6)', 'Calibri', 10)
    p2 = cell.add_paragraph()
    run_styled(p2, '\u2022 HP ProDesk 2, AMD Ryzen 5  \u2022 16 GB RAM  \u2022 512 GB SSD  \u2022 HP 21.5" LED Monitor', 'Calibri', 9.5)
    p3 = cell.add_paragraph()
    p3.paragraph_format.space_before = Pt(4)
    run_styled(p3, 'Total Investment: \u20b923,12,800', 'Cambria', 12, bold=True, color=NAVY)

highlight_box(equip_content)
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 7 — PLACEMENTS
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 7: Placements")
section_header("Overview of Academic Programs", "SECTION 07 \u2014 Campus Placements \u00b7 Internships \u00b7 Higher Studies")

doc.add_heading("Campus Placements 2025\u201326", level=2)
stat_cards([
    ("98.21%", "Placement Rate", "green"),
    ("\u20b933 LPA", "Highest CTC", "navy"),
    ("\u20b910.60 LPA", "Average CTC", "gold"),
    ("110", "Students Placed", "navy"),
])

data_table(
    ["Metric", "Value"],
    [["Total Students", "155"], ["Opted for Placements", "118"], ["Placed", "110"],
     ["Placement Percentage", "98.21%"], ["Highest CTC", "\u20b933.00 LPA"],
     ["Lowest CTC", "\u20b93.60 LPA"], ["Average CTC", "\u20b910.60 LPA"], ["Median CTC", "\u20b99.00 LPA"]]
)

doc.add_heading("Top Recruiting Companies", level=3)
data_table(
    ["Company", "Offers"],
    [["Barclays", "21"], ["Nomura", "6"], ["BlackRock", "5"], ["IKS Health", "5"],
     ["ICICI Lombard", "5"], ["General Mills", "4"], ["Baker Hughes", "4"],
     ["We Matter", "4"], ["PhonePe", "3"], ["JP Morgan & Chase", "3"],
     ["HSBC", "3"], ["Loginext", "3"], ["Wissen Technology", "2"],
     ["UBS", "2"], ["Morgan Stanley", "1"], ["BrowserStack", "1"]],
    compact=True
)

doc.add_heading("Our Recruiters", level=3)
elite = {"Barclays", "Nomura", "BlackRock", "JP Morgan", "Morgan Stanley", "HSBC", "PhonePe"}
company_tags(["Barclays", "Nomura", "BlackRock", "JP Morgan", "Morgan Stanley", "HSBC", "PhonePe",
    "General Mills", "Baker Hughes", "ICICI Lombard", "IKS Health", "Wissen Technology",
    "IDfy", "Russell Investments", "Zeta", "UBS", "BSE", "BrowserStack", "Nouryon",
    "Revvity", "ISS Stoxx", "EightGen AI", "IDFC", "WorkIndia", "Auxo AI",
    "Hyperverge", "Loginext", "We Matter", "Mahindra Finance"], elite)

doc.add_page_break()

doc.add_heading("Higher Studies \u2014 Batch 2025\u201326", level=2)
stat_cards([("24", "IIT Bombay", "navy"), ("4", "IIT Patna", "gold"), ("32", "Total Higher Studies", "green")])

data_table(
    ["Institution", "No. of Students"],
    [["Indian Institute of Technology, Bombay", "24"], ["IIT Patna", "4"],
     ["IIT Madras", "2"], ["IIT Guwahati", "2"], ["SPJIMR", "2"], ["DRDO", "1"]]
)

doc.add_heading("Industry Internships", level=2)
body("Every student completes a mandatory industry or research internship at organisations including Amazon, BlackRock, Baker Hughes, HSBC, JP Morgan, General Mills, Morgan Stanley, PhonePe, Nomura, BrowserStack, and 50+ more.")

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 8 — FACULTY
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 8: Faculty & Staff")
section_header("Faculty and Staff", "SECTION 08")

body_rich([
    ("The Department is supported by ", False),
    ("30 faculty members", True, False, NAVY),
    (" \u2014 12 with doctoral degrees, 6 pursuing Ph.D. \u2014 and 6 technical and supporting staff.", False),
])

add_image(os.path.join(BASE, 'Faclty group pic.JPG'), width=Inches(5.5))
caption("Faculty Group \u2014 Department of Computer Engineering, 2025\u201326")

doc.add_heading("Faculty with Doctoral Degrees", level=3)
data_table(["Sr.", "Name", "Designation"],
    [["1", "Dr. Sudhir N. Dhage", "Professor & HOD"], ["2", "Dr. Prasenjit Bhavathankar", "Professor"],
     ["3", "Dr. Surekha Dholay", "Professor"], ["4", "Dr. Anant V. Nimkar", "Professor"],
     ["5", "Dr. Nataasha Raul", "Professor"], ["6", "Dr. Kailas K. Devadkar", "Professor"],
     ["7", "Dr. Pramod J. Bide", "Professor"], ["8", "Dr. Disha Sail", "Professor"],
     ["9", "Dr. Jignesh Sisodia", "Professor"], ["10", "Dr. Anuj Tawari", "Professor"],
     ["11", "Dr. Sanjuktarani Jena", "Professor"]], compact=True)

doc.add_heading("Assistant Professors", level=3)
data_table(["Sr.", "Name"], [
    [str(i+1), n] for i, n in enumerate([
        "Prof. Anand Godbole", "Prof. Reeta Koshy", "Prof. Sunil Ghane", "Prof. Abhijeet Salunke",
        "Prof. Jyoti Ramteke", "Prof. Swapnali Kurhade", "Prof. Govind Gaundalkar", "Prof. Soni Bhambar",
        "Prof. Yogita Patil", "Prof. Jotsna Bhagat", "Prof. Isha Sawalkar", "Prof. Azman Qureshi",
        "Prof. Falak Malik", "Prof. Ameya Nile", "Prof. Amey Morye", "Prof. Sudeep Mali",
        "Ms. Aishwarya Nalawade", "Mr. Anas Ansari", "Mr. Udit Kangad", "Prof. Vishwas Badhe",
    ])
], compact=True)

doc.add_heading("Technical & Supporting Staff", level=3)
data_table(["Category", "Name"], [
    ["Technical Assistant", "Mr. Rajendra Nalawade"], ["Technical Assistant", "Mr. Nandakumar Kinare"],
    ["Technical Assistant", "Mrs. Rupali Bhosale"], ["Lab Attendant", "Mr. Mangesh Birje"],
    ["Lab Attendant", "Mr. Vishwanath Kamblekar"], ["Lab Attendant", "Mr. Prathmesh Bhonde"],
], compact=True)

placeholder("ADD INDIVIDUAL FACULTY PHOTOGRAPHS IF AVAILABLE")
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 9 & 10 — RESPONSIBILITIES & ADVISORY BOARD
# ═════════════════════════════════════════════════════════════
print("  \u2192 Sections 9 & 10: Responsibilities & Advisory Board")
section_header("Additional Responsibilities & Advisory Board", "SECTIONS 09 & 10")

doc.add_heading("Additional Responsibilities to the Faculty", level=2)
data_table(["Faculty Member", "Responsibility"], [
    ["Dr. Sudhir N. Dhage", "Head of Department, Chairman BoS"],
    ["Dr. Prasenjit Bhavathankar", "NAAC Coordinator, NBA Documentation"],
    ["Dr. Surekha Dholay", "Research Coordinator, TPC Reviewer"],
    ["Dr. Anant V. Nimkar", "NBA Accreditation Coordinator"],
    ["Dr. Kailas Devadkar", "Patent & IPR Cell Coordinator"],
    ["Prof. Anand Godbole", "NPTEL Coordinator"],
    ["Prof. Reeta Koshy", "DAPC Member, Result Analysis"],
    ["Prof. Sunil Ghane", "Project Coordinator (BE), IQAC Member"],
    ["Prof. Swapnali Kurhade", "Placement Coordinator, Slow-Fast Learner"],
    ["Prof. Jyoti Ramteke", "Co-Teaching Coordinator"],
    ["Prof. Govind Gaundalkar", "Exam Cell, LMS Moodle Admin"],
    ["Prof. Yogita Patil", "Liberal Learning Course Coordinator"],
    ["Prof. Abhijeet Salunke", "Lab Coordinator, Assessment Tools"],
])

placeholder("ADD / EDIT ADDITIONAL RESPONSIBILITIES")
gold_rule()

doc.add_heading("Advisory Board & Board of Studies", level=2)
body("The Department maintains an active Advisory Board of eminent academicians and industry leaders. The BoS reviews curriculum for industry alignment. The DAPC monitors academic progress and ensures quality assessment (ISE 20%, MSE 20%, ESE 100%).")

placeholder("ADD ADVISORY BOARD MEMBERS LIST")
placeholder("ADD BOARD OF STUDIES MEMBERS LIST")
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 11 — RESEARCH
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 11: Research Publications")
section_header("Consultancy, Grants, Patents & Research Publications", "SECTION 11")

stat_cards([("15", "Journals", "navy"), ("25", "Conferences", "gold"),
            ("9", "Scopus", "navy"), ("3", "SCIE", "green"), ("1", "Patent", "gold")])

doc.add_heading("Patent Filed", level=3)
def patent_content(cell):
    p = cell.paragraphs[0]
    run_styled(p, 'SOUND2SKETCH: An AI-Powered Interactive Voice-Enabled Drawing Tool', 'Cambria', 11, bold=True, color=NAVY)
    p2 = cell.add_paragraph()
    run_styled(p2, 'Inventors: ', 'Calibri', 9.5, bold=True)
    run_styled(p2, 'Dr. Kailas Devadkar, Satyam Jaiswal, Siddhant Rao, Dr. Jignesh Sisodia, Komal Tarachandani', 'Calibri', 9.5)
    p3 = cell.add_paragraph()
    run_styled(p3, 'Application No: ', 'Calibri', 9.5, bold=True)
    run_styled(p3, '20 2025 104 311  |  ', 'Calibri', 9.5)
    run_styled(p3, 'Filed: ', 'Calibri', 9.5, bold=True)
    run_styled(p3, '25 July 2025  |  ', 'Calibri', 9.5)
    run_styled(p3, 'Published: ', 'Calibri', 9.5, bold=True)
    run_styled(p3, '20 August 2025', 'Calibri', 9.5)

highlight_box(patent_content)

doc.add_heading("Faculty Citation Metrics", level=3)
data_table(["Sr.", "Faculty", "Total Citations", "Citations (2026)"], [
    ["1", "Dr. Sudhir N. Dhage", "309", "111"], ["2", "Dr. Kailas K. Devadkar", "118", "37"],
    ["3", "Dr. Anant V. Nimkar", "92", "22"], ["4", "Mr. Sunil B. Ghane", "89", "24"],
    ["5", "Dr. Pramod J. Bide", "85", "11"], ["6", "Dr. Nataasha N. Raul", "64", "14"],
    ["7", "Dr. Surekha Dholay", "62", "10"], ["8", "Dr. Prasenjit Bhavathankar", "45", "10"],
    ["9", "Mrs. Reeta Koshy", "40", "9"], ["10", "Prof. Anand Godbole", "39", "12"],
    ["11", "Dr. Jignesh Sisodia", "28", "11"], ["12", "Mrs. Jyoti Ramteke", "28", "7"],
    ["13", "Dr. Disha Sail", "28", "1"], ["14", "Mrs. Swapnali Kurhade", "27", "6"],
    ["15", "Prof. Abhijeet Salunke", "21", "7"],
], compact=True)

doc.add_page_break()

doc.add_heading("Select Journal Publications (2025\u201326)", level=2)
data_table(["Sr.", "Title", "Author(s)", "Journal", "Index"], [
    ["1", "Emoji-Free Sarcasm Detection in Code-Mixed Text", "Falak Malik et al.", "IEEE Trans. Affective Computing", "Scopus"],
    ["2", "Physics-based Forgery Detection using RBF-SVM", "Dr. Anant Nimkar", "Signal, Image & Video (Springer)", "SCIE"],
    ["3", "Hybrid MoE & VLM for Plant Disease Detection", "Dr. Disha Sail", "AGH Computer Science", "SCIE"],
    ["4", "Multimodal Assistant for Visually Impaired", "Mrs. Reeta Koshy", "Int'l J. Technology (NTU)", "SCIE"],
    ["5", "Data-Driven Agro-Climatic Zoning of India", "Dr. K. Devadkar", "Journal of Technology", "Scopus"],
    ["6", "Multimodal Anomaly Detection for Police Complaints", "Dr. K. Devadkar, Dr. J. Sisodia", "Journal of Technology", "Scopus"],
    ["7", "AI-Driven Renewable Energy Solutions", "Dr. Nataasha Raul", "Journal of Technology", "Scopus"],
], compact=True)

doc.add_heading("Select Conference Publications (2025\u201326)", level=2)
data_table(["Sr.", "Title", "Author(s)", "Conference"], [
    ["1", "Physics-Based AR for Home Furnishing", "Dr. Anant Nimkar", "IEEE IATMSI 2026"],
    ["2", "ML & GenAI for Vulnerability Detection", "Dr. Anant Nimkar", "IEEE IATMSI 2026"],
    ["3", "Cryptographic API Misuse Detection (BERT)", "Dr. Anant Nimkar", "IEEE ICITIIT 2026"],
    ["4", "CyberVigilance: DL for Cyber Attack Forecasting", "Swapnali Kurhade", "IEEE INDISCON 2025"],
    ["5", "LexChain: Blockchain for Electronic Evidence", "Dr. P. Bhavathankar", "IEEE ISDFS 2025"],
    ["6", "Graph Neural Approaches to Credit Scoring", "Dr. J. Sisodia et al.", "IEEE APCIT 2025"],
    ["7", "Beyond Hate Speech: Cyberbullying Detection", "Sunil Ghane", "IEEE PuneCon 2025"],
    ["8", "KubeTranscode: Distributed DASH Transcoding", "Dr. Pramod Bide", "IEEE GIEST 2025"],
    ["9", "Flood Risk Prediction using ML", "Falak Malik", "IEEE INDISCON 2025"],
    ["10", "Tutor-Aware Essay Scoring (Knowledge Graph)", "Dr. K. Devadkar", "IEEE ICMI 2026"],
], compact=True)
body("Showing 10 of 25 conference publications.", italic=True)
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 12 — ACHIEVEMENTS
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 12: Achievements")
section_header("Faculty and Students Achievements", "SECTION 12")

doc.add_heading("Faculty Achievements (2025\u201326)", level=2)
data_table(["Faculty", "Achievement"], [
    ["Prof. Anand Godbole", "NPTEL Discipline Star Award"],
    ["Dr. Pramod Bide", "Best Innovative Teaching Award \u2014 L&T Edutech"],
    ["Dr. Jignesh Sisodia", "Best Software Project Award \u2014 Cantilever"],
    ["Prof. Govind Gaundalkar", "Best Question Paper \u2014 Digital Systems"],
    ["Dr. J. Sisodia + Prof. A. Salunke", "Best Question Paper \u2014 CCN"],
    ["Dr. Surekha Dholay", "TPC Reviewer"],
])

gold_rule()

doc.add_heading("Student Achievements \u2014 Highlights", level=2)

achievements = [
    ("EY Techathon 6.0 \u2014 National Winner \U0001f3c6", "Soham Ghorpade, Dikshant Karande, Dwayne Fernandes", "Ernst & Young (EY)", "WINNER", GOLD_HEX),
    ("ICPC Asia Amritapuri \u2014 Regionalist", "Raj Kale, Abhimanyu Kadhane, Adityavardhan Iyengar", "ICPC Foundation", "REGIONALIST", NAVY_HEX),
    ("GenAI Exchange Hackathon \u2014 3rd Runner Up", "Daksh Chetan Jain", "Google Cloud", "3RD RUNNER UP", NAVY_HEX),
    ("Ethereum Bug Bounty \u2014 Hall of Fame", "Madhav Shah", "Ethereum Foundation", "HALL OF FAME", GREEN_HEX),
    ("Zoho Bug Bounty \u2014 Hall of Fame", "Madhav Shah", "Zoho", "HALL OF FAME", GREEN_HEX),
    ("HackFusion International \u2014 Winner", "Vir Bhalani, Daksh Jain, Vishwesh Nair, Jill Mistry", "IEEE SPIT & UMIT", "WINNER", GOLD_HEX),
    ("IIT Bombay AI Eduthon \u2014 Winner", "Daksh Jain", "Cograd, IIT Bombay", "WINNER", GOLD_HEX),
    ("Chaitanya Case Competition \u2014 1st Place", "Sanjana Pillai, Aditya Khabiya", "IIM Indore", "1ST PLACE", GOLD_HEX),
    ("SPIT Hackathon \u2014 Domain Winner", "Vir Bhalani, Oren Coelho, Piyush Rathi, Daksh Jain", "CSI SPIT", "DOMAIN WINNER", NAVY_HEX),
    ("Marketing Showdown \u2014 Winner", "Soham Kher, Aniruddh Iyer", "IIT Bombay SJMSOM", "WINNER", GOLD_HEX),
]
for event, students, org, result, accent in achievements:
    achievement_card(event, students, org, result, accent)

placeholder("ADD ADDITIONAL STUDENT ACHIEVEMENTS, SPORTS, DEBATE RESULTS IF AVAILABLE")
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 13 — PROJECTS
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 13: Projects")
section_header("Undergraduate Projects, M.Tech & Ph.D.", "SECTION 13")

doc.add_heading("B.Tech Major Projects \u2014 AY 2025\u201326", level=2)
data_table(["Gr.", "Project Title", "Students", "Guide", "Domain"], [
    ["1", "AI-driven Financial Portfolio Management", "Atharva Bambere et al.", "Dr. B.N. Chaudhari", "ML"],
    ["2", "Financial Health Evaluation & Wealth Advisory", "Maureen Miranda et al.", "Dr. Sudhir Dhage", "ML"],
    ["3", "Deep Models for Alzheimer\u2019s Detection (MRI)", "Shardul Mahindrakar et al.", "Dr. Sudhir Dhage", "DL"],
    ["4", "AI-Powered Learning Platform", "Kanishka Amritkar et al.", "Dr. Sudhir Dhage", "AI"],
    ["5", "Dynamic Route Optimization", "Sneha Sambare et al.", "Dr. Surekha Dholay", "ML"],
    ["6", "Deepfake Detection (XceptionNet & DCT)", "Shrinivas Waney et al.", "Dr. Surekha Dholay", "DL"],
    ["7", "FundChain: Blockchain Fundraising", "Manish Jadhav et al.", "Dr. P. Bhavathankar", "Blockchain"],
    ["8", "BharatNFT: Digital Asset Marketplace", "Shesh Shrikant et al.", "Dr. P. Bhavathankar", "Blockchain"],
    ["9", "Mindease: Stress Detection System", "Harsh Shah et al.", "Prof. Anand Godbole", "AI"],
    ["10", "Injury Detection for Football (DL)", "Daksh Jain et al.", "Prof. Anand Godbole", "DL"],
    ["11", "AI Recruitment & Career Platform", "Aayush Dubey et al.", "Prof. Anand Godbole", "AI"],
], compact=True)
body("Showing select projects. Total: 35+ project groups.", italic=True)

doc.add_page_break()
doc.add_heading("M.Tech Dissertation Projects \u2014 AY 2025\u201326", level=2)
data_table(["Sr.", "Student", "Dissertation Title", "Mentor"], [
    ["1", "Ameya Nile", "ML-Based Food Waste Reduction & Restaurant Mgmt", "Dr. Sudhir Dhage"],
    ["2", "Tushar Barkur", "Explainable Cross-Domain User Preference Modeling", "Dr. Jignesh Sisodia"],
    ["3", "Vaishnavi Dalvi", "ClaimSage: Agentic RAG for Insurance Claims", "Prof. Reeta Koshy"],
    ["4", "Vedang Deshpande", "RL for Energy Optimization in Smart Grids", "Dr. Aarti Karande"],
    ["5", "Suyog Dhande", "Explainable AI for Legal Case Prioritization", "Dr. D.R. Kalbande"],
    ["6", "Om Jadhao", "Image QA for Ethical Education", "Prof. Jyoti Ramteke"],
    ["7", "Gauri Korgaonkar", "Emotion-Aware Zero Trust Access Control", "Dr. Renuka Pawar"],
    ["8", "Sudarshan Lagade", "Water Demand Prediction in India", "Prof. Abhijit Salunke"],
    ["9", "Durgesh Mahant", "Smart Navigation with Hazard Detection", "Dr. Nataasha Raul"],
    ["10", "Kiran Maharana", "Spatiotemporal GNNs for Water Contamination", "Dr. P.B. Bhavathankar"],
    ["11", "Falak Malik", "Emoji-Free Sarcasm Detection (Code-Mixed)", "Dr. Pramod Bide"],
    ["12", "Amey Morye", "Yoga Mitra: Multilingual Yoga Framework", "Dr. Sujata Kulkarni"],
    ["13", "Devang Parab", "AI-Powered Mental Health Prediction", "Dr. K.K. Devadkar"],
    ["14", "Fazil Shaikh", "Decomposing Recall in Small Language Models", "Prof. S. Choudhary"],
    ["15", "Kshitija Sole", "AgriXplain: Ensemble Crop Yield Forecasting", "Prof. Sunil Ghane"],
], compact=True)

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 14 — INDUSTRY INTERACTION
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 14: Industry Interaction")
section_header("Industry\u2013Institute Interaction", "SECTION 14")

body("The Department actively engages with industry through MOUs, co-teaching, workshops, and collaborative programmes.")

doc.add_heading("Key Collaborations", level=3)
bullet("Banking Technology \u2014 Barclays", bold_prefix="MDM: ")
bullet("Management \u2014 SPJIMR", bold_prefix="MDM: ")
bullet("NVIDIA Deep Learning Institute (DLI) cohort", bold_prefix="Skill Programme: ")
bullet("Program Elective by Industry \u2014 Vizuara", bold_prefix="Industry Elective: ")

doc.add_heading("Co-Teaching", level=3)
body_rich([("AIML Co-Teaching: ", True, False, NAVY),
    ("Dr. Maheshkumar H. Kolekar, IIT Patna (13 Oct 2024), arranged by Prof. Swapnali Kurhade.", False)])

placeholder("ADD COMPLETE LIST OF MOUs")
placeholder("ADD WORKSHOPS / STTP / FDP / MOOC")
placeholder("ADD EXPERT TALKS BY FACULTY")
placeholder("ADD STUDENT INDUSTRY VISITS")
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 15 — ALUMNI
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 15: Alumni")
section_header("Distinguished Alumni of the Department", "SECTION 15")

body("The Department takes pride in its alumni who lead across technology, finance, and entrepreneurship. Our alumni network contributes through mentoring, guest lectures, co-teaching, and recruitment.")

doc.add_heading("Alumni Mentoring Programme (2025\u201326)", level=3)
add_image(os.path.join(IMG, 'ptm_slide15_Google_Shape_269_p91.png'), Inches(5.0))
caption("Alumni Mentoring Programme 2025\u201326")

doc.add_heading("Distinguished Alumni", level=3)
add_image(os.path.join(IMG, 'ptm_slide16_Google_Shape_279_p98.png'), Inches(5.0))
caption("Distinguished Alumni of the Department")

placeholder("ADD ALUMNI PROFILES \u2014 Name, Batch, Current Role, Company, Photo")
placeholder("ADD ALUMNI ENTREPRENEURS LIST")
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 16 — ACTIVITIES
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 16: Activities")
section_header("Departmental Activities", "SECTION 16 \u2014 FACE \u00b7 CSI \u00b7 IEEE \u00b7 Events")

doc.add_heading("Parent-Teacher Meeting", level=2)
add_image(os.path.join(BASE, 'PTM.JPG'), Inches(5.0))
caption("Parent-Teacher Meeting 2025\u201326 (30 April 2026)")

doc.add_heading("CSI \u2014 Computer Society of India", level=2)

doc.add_heading("CodeHousie (16 Oct 2025)", level=3)
body("An innovative format integrating coding challenges with the traditional Housie game, held in the fourth-floor computer laboratories.")

doc.add_heading("Pixel Perfect (14 Feb 2026)", level=3)
body("A frontend-focused hackathon by CSI-SPIT \u00d7 ACSES SPIT, challenging students to build pixel-perfect, responsive web interfaces.")

doc.add_heading("SPIT Hackathon 2026 \u2014 10th Edition \U0001f3c6", level=3)
body_rich([
    ("The flagship 24-hour national hackathon (21\u201322 Feb 2026). ", False),
    ("3,316 registrations", True, False, NAVY), (", ", False),
    ("78 finalist teams", True, False, NAVY), (", prize pool of ", False),
    ("\u20b92.5 Lakhs", True, False, NAVY),
    (". Tracks: AIML, Web Dev, App Dev, Blockchain.", False),
])

gold_rule()

doc.add_heading("IEEE Computer Society \u2014 SPIT", level=2)
bullet("Brain Drain \u2014 Technical quiz across CS domains")
bullet("Research Ideation \u2014 Student research proposal programme")
bullet("Quantum Computing Conclave \u2014 Quantum gates, qubits, algorithms")
bullet("AIRAVAT 3.0 \u2014 AI/ML hands-on workshop & competition")

doc.add_heading("FACE \u2014 Codefolio", level=2)
body("Competitive programming and portfolio-building event for placement and contest preparation.")

doc.add_heading("Liberal Learning Courses (LLC)", level=2)
body("Activity-based courses: Yoga, Painting, Kathak, First Aid, NGO Work, Bhagavad Gita, Theatre, Rangoli, Photography, Music, Martial Arts.")

for img in ['ptm_slide25_Google_Shape_478_p96.jpg', 'ptm_slide25_Google_Shape_479_p96.jpg']:
    add_image(os.path.join(IMG, img), Inches(3.5))

placeholder("ADD MORE EVENT PHOTOS AND DESCRIPTIONS")
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 17 — STUDENT SUPPORT
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 17: Student Support")
section_header("Student Support & Innovative Teaching", "SECTION 17")

doc.add_heading("Slow & Fast Learner Support", level=2)
bullet("Slow Learners: Remedial assignments in OS, AISC, AIML; tutorial sessions; mentoring \u2014 60 hours", bold_prefix="")
bullet("Fast Learners: Disk Scheduling Simulator (~54 hrs), IEEE paper review, hyperparameter tuning \u2014 57 hours", bold_prefix="")

thin_rule()

doc.add_heading("Innovative Teaching\u2013Learning Practices", level=2)
data_table(["Faculty", "Innovation", "Subject", "Class"], [
    ["Prof. Jyoti Ramteke", "Fun Quiz using Mentimeter", "AISC", "TE Comp C&D"],
    ["Prof. Reeta Koshy", "Flipped Classroom \u2014 Group Discussion", "DevOps", "TE Comp AB"],
    ["Prof. Swapnali Kurhade", "Mentimeter, Think-Pair-Share, Open Book", "AIIA", "Final Year"],
    ["Prof. Abhijeet Salunke", "Card Trick for cryptography concepts", "CNS", "TE Comp"],
    ["Prof. Anand Godbole", "The 5-Card Trick for discrete structures", "DSGT", "SE"],
    ["Dr. Pramod Bide", "QR Code scanning via mobile for HPC", "HPC", "BE"],
])

doc.add_heading("Alumni Connect & Co-Teaching", level=3)
body_rich([("AIML Co-Teaching: ", True, False, NAVY),
    ("Dr. Maheshkumar H. Kolekar, IIT Patna (13 Oct 2024).", False)])

doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  SECTION 18 — SWOC
# ═════════════════════════════════════════════════════════════
print("  \u2192 Section 18: SWOC")
section_header("Departmental SWOC Analysis", "SECTION 18")

body("A candid assessment of the department\u2019s current position and future trajectory.")

swoc_box("\u2666 Strengths", [
    "98.21% placement with top-tier companies", "NBA Accredited 4 consecutive cycles",
    "40 publications (15 journals + 25 conferences), 9 Scopus, 3 SCIE",
    "12 PhD faculty, 6 pursuing", "NEP 2020 Pi-Model with 57 pathways",
    "Active CSI, IEEE, FACE, ACSES", "24 students placed at IIT Bombay",
    "SPIT Hackathon: 3,316 registrations",
], GREEN_BG, GREEN_HEX)

swoc_box("\u2666 Weaknesses", [
    "Limited industry-sponsored research grants", "Some contractual faculty positions",
    "Need for Quantum/Robotics research labs", "Limited international collaborations",
    "Doctoral programme needs more scholars",
], RED_BG, RED_HEX)

swoc_box("\u2666 Opportunities", [
    "Growing AI/ML, Cybersecurity demand", "NEP 2020 programme design flexibility",
    "Mumbai startup ecosystem for incubation", "Digital India research funding",
    "MDM expansion with industry leaders", "Patent portfolio growth",
], BLUE_SOFT, NAVY_HEX)

swoc_box("\u2666 Challenges", [
    "Rapid tech change demands curriculum updates", "Faculty retention in competitive market",
    "Balancing research with teaching loads", "Rising infrastructure costs",
    "Competition from coding academies",
], ORANGE_BG, 'E65100')

placeholder("EDIT / ADD SWOC POINTS AS NEEDED")
doc.add_page_break()


# ═════════════════════════════════════════════════════════════
#  MEET THE TEAM
# ═════════════════════════════════════════════════════════════
print("  \u2192 Meet the Team")

# Centered credits page
for _ in range(5):
    add_spacer(10)

gold_rule()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'MEET THE TEAM', 'Cambria', 24, bold=True, color=NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'This magazine was conceptualised, curated, and crafted by the editorial team\nof the Department of Computer Engineering, S.P.I.T., Mumbai.', 'Calibri', 10, color=GREY)

add_spacer(10)

placeholder("ADD EDITORIAL TEAM \u2014 Name, Role, Photo")
placeholder("ADD FACULTY ADVISOR NAME AND PHOTO")

add_spacer(12)
gold_rule()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, "Bharatiya Vidya Bhavan\u2019s", 'Cambria', 10, italic=True, color=NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'SARDAR PATEL INSTITUTE OF TECHNOLOGY', 'Cambria', 14, bold=True, color=NAVY)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, 'Munshi Nagar, Andheri (West), Mumbai \u2014 400 058\nhttps://comp.spit.ac.in/', 'Calibri', 8.5, color=GREY)

add_spacer(12)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_styled(p, '\u00a9 2026 Department of Computer Engineering, S.P.I.T. \u2014 Annual Magazine 2025\u201326', 'Calibri', 7.5, color=LIGHT_GREY)


# ═════════════════════════════════════════════════════════════
#  SAVE
# ═════════════════════════════════════════════════════════════
print(f"\nSaving to: {OUT}")
doc.save(OUT)
sz = os.path.getsize(OUT) / 1024
print(f"\u2705 Premium Magazine DOCX created successfully!")
print(f"   File: {OUT}")
print(f"   Size: {sz:.0f} KB")
