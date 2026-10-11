"""
verified_builder_part2.py - Sections 5 to 7: Curriculum, Labs & Facilities, Academic Programs, Placements & Internships
"""

import os, json
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from magazine_core import (
    NAVY_HEX, GOLD_HEX, GOLD_LIGHT_HEX, GOLD_BG_HEX, LIGHT_BG,
    NAVY, GOLD, GOLD_DARK, WHITE, BLACK, GREY,
    run_styled, add_spacer, section_header, sub_section_header,
    stat_cards, callout_box, data_table, body, bullet, add_image_safe
)

SCRATCH = r'C:\Users\Nirupam\.gemini\antigravity-ide\brain\d84a1836-479b-4390-be26-b801a6bee109\scratch'
MEDIA_DIR = os.path.join(SCRATCH, 'baseline_media')

def load_docx_data():
    with open(os.path.join(SCRATCH, 'docx_data.json'), 'r', encoding='utf-8') as f:
        return json.load(f)

def build_section_05(doc):
    section_header(doc, "About the Department & Curriculum Architecture", "SECTION 05 — ACADEMIC REVOLUTION")

    # Program History Table
    sub_section_header(doc, "Program History & Approved Intakes", "5.1")
    hist_headers = ["Program Offered", "Branch", "Year of Estd.", "Sanctioned Intake"]
    hist_rows = [
        ["B.Tech.", "Computer Engineering", "1995", "60"],
        ["B.Tech.", "Computer Engineering", "2021", "120"],
        ["B.Tech.", "Computer Engineering", "2023", "240"],
        ["M.Tech.", "Computer Engineering", "2011", "18"],
        ["Ph.D.", "Computer Engineering", "2012", "10"],
        ["Ph.D.", "Computer Engineering", "2016", "20"],
        ["Ph.D.", "Computer Engineering", "2019", "30"]
    ]
    data_table(doc, hist_headers, hist_rows, col_widths=[1.5, 2.5, 1.5, 1.5])

    body(doc, "Accreditation Milestones (External Audit PDF Page 2):", bold=True, color=NAVY)
    bullet(doc, "NBA Accredited in 2013–14 (for 2 years), 2016–17 (AY 2017–18 to 2019–20), 2021–22 (AY 2021–22 to 2023–24), and re-accredited in 2024–25 (valid up to 30 June 2027).")
    bullet(doc, "NAAC 'A' Grade Accreditation awarded on June 28, 2025 (Valid until June 2027).")

    add_spacer(doc, 6)

    # Narrative from 5. About the Department.docx
    sub_section_header(doc, "Redefining Engineering Education: The New-Age Curriculum at S.P.I.T.", "5.2")
    data = load_docx_data()
    paragraphs = data['5. About the Department.docx']['paragraphs']
    
    for p_text in paragraphs:
        pt = p_text.strip()
        if not pt:
            continue
        # Check if heading-like
        if pt in [
            "Redefining Engineering Education: The New-Age Curriculum at S.P.I.T.",
            "From Learning Technology to Creating Impact",
            "A Curriculum Built Around Choice",
            "Building Strong Foundations",
            "From Core Computing to Emerging Technologies",
            "A Strong Emphasis on Artificial Intelligence and Emerging Domains",
            "Beyond Silos: Multidisciplinary Learning",
            "Cultivating the Research Mindset: Honours by Research",
            "Real-World Immersion: Industry Internships and Experiential Learning",
            "Developing the Complete Professional: Liberal Learning and Ethics",
            "A Culture of Innovation and Entrepreneurship",
            "Graduating Future-Ready Engineers",
            "Conclusion: Education for Tomorrow"
        ]:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.keep_with_next = True
            run_styled(p, pt, 'Cambria', 11.5, bold=True, color=NAVY)
        else:
            body(doc, pt)

    doc.add_page_break()

def build_section_06(doc):
    section_header(doc, "Laboratory Facilities, Smart Classrooms & Equipment", "SECTION 06 — INFRASTRUCTURE EXCELLENCE")

    sub_section_header(doc, "Department Laboratories Roster", "6.1")
    body(doc, "The Department houses 15 specialized, ICT-enabled laboratories to support undergraduate, postgraduate, and doctoral computing instruction and advanced research:")

    labs_headers = ["Sr. No.", "Laboratory Name", "Location / Room No."]
    labs_rows = [
        ["1", "Networking Laboratory", "Room 603"],
        ["2", "Algorithms Laboratory", "Sixth Floor"],
        ["3", "Post Graduate (PG) Laboratory", "Sixth Floor"],
        ["4", "Distributed Computing Laboratory", "Room 604"],
        ["5", "Database Laboratory", "Room 606"],
        ["6", "Data Analytics Laboratory", "Room 606"],
        ["7", "Software Engineering Laboratory", "Room 607 B"],
        ["8", "System Programming Laboratory", "Room 608"],
        ["9", "Machine Learning Laboratory", "Room 702"],
        ["10", "Artificial Intelligence Laboratory", "Seventh Floor"],
        ["11", "Soft Computing Laboratory", "Seventh Floor"],
        ["12", "System Security Laboratory", "Room 703"],
        ["13", "Cyber Security Laboratory", "Seventh Floor"],
        ["14", "Operating Systems Laboratory", "Room 507"],
        ["15", "Project Laboratory", "Room 509"]
    ]
    data_table(doc, labs_headers, labs_rows, col_widths=[0.8, 4.2, 2.0])

    sub_section_header(doc, "Smart Classrooms", "6.2")
    body(doc, "Department lectures and interactive seminars are conducted in high-tech, smart-equipped classrooms including Rooms 203, 408, 508, 601, 605, 609, and 002. All classrooms feature acoustic optimization, high-lumen projection systems, and campus-wide high-speed Wi-Fi.")

    # Lab Photographs
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image2.jpg"), width=Inches(3.2), caption_text="Room 604 — Distributed Computing Laboratory")
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image3.jpg"), width=Inches(3.2), caption_text="Room 606 — Data Analytics Laboratory")
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image4.jpg"), width=Inches(3.2), caption_text="Room 507 — Operating Systems Laboratory")
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image5.jpg"), width=Inches(3.2), caption_text="Room 509 — Project Laboratory")

    sub_section_header(doc, "Major Equipment & Workstations Added in Laboratories during 2025–26", "6.3")
    body(doc, "As verified in the External Academic Audit (Audit PDF Page 21), the department continuously upgrades computing infrastructure. In AY 2025–26, 80 high-performance HP ProDesk computing workstations were deployed:")

    eq_headers = ["Sr. No.", "Workstation Configuration", "Qty", "Cost (₹)", "Room No. / Lab"]
    eq_rows = [
        ["1", "HP PRODESK 2 GLA, Desktop-D2YL7AT, R5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor", "20", "11,56,400/-", "Lab 5 (Room No. 606)"],
        ["2", "HP PRODESK 2 GLA, Desktop-D2YL7AT, R5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor", "20", "11,56,400/-", "Lab 6 (Room No. 606)"],
        ["3", "HP PRODESK 2 GLA, Desktop-D2YL7AT, Core i5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor", "20", "13,15,794/-", "Lab 4 (Room No. 604)"],
        ["4", "HP PRODESK 2 GLA, Desktop-D2YL7AT, Core i5, 16 GB RAM, 512 GB SSD, HP 21.5\" LED Monitor", "20", "13,15,794/-", "Lab 8 (Room No. 608)"],
        ["TOTAL", "Total Workstations Added across 4 Laboratories", "80", "49,44,388/-", "Labs 4, 5, 6, 8"]
    ]
    data_table(doc, eq_headers, eq_rows, col_widths=[0.6, 3.4, 0.6, 1.2, 1.2], compact=True)

    doc.add_page_break()

def build_section_07(doc):
    section_header(doc, "Overview of Academic Programs, Placements & Internships", "SECTION 07 — CAREER TRAJECTORY & OUTCOMES")

    stat_cards(doc, [
        ("98.21%", "Placement Rate"),
        ("₹33.00 LPA", "Highest CTC"),
        ("₹10.60 LPA", "Average CTC"),
        ("₹9.00 LPA", "Median CTC"),
        ("110", "Students Placed")
    ])

    sub_section_header(doc, "Comprehensive Placement Summary (B.Tech Computer Engineering)", "7.1")
    body(doc, "Official placement audit metrics for the 2022–2026 Batch (2025–26 Passout) from 7. Overview of Academic Programs.docx:")

    summary_headers = ["Company Category", "Offers", "Student Metrics", "Count", "Salary Metrics", "CTC (LPA)"]
    summary_rows = [
        ["Elite (≥ 15 LPA)", "21", "Total Enrolled Students", "155", "Highest CTC", "₹33.00 LPA"],
        ["Super Dream (10 – 14.9 LPA)", "70", "Opted for Placements", "118", "Lowest CTC", "₹3.60 LPA"],
        ["Dream (6 – 9.9 LPA)", "10", "Students Placed", "110", "Average CTC", "₹10.60 LPA"],
        ["Normal (< 6 LPA)", "9", "Ineligible Students", "6", "Median CTC", "₹9.00 LPA"],
        ["Ineligible Categories", "7", "Placement Percentage", "98.21%", "—", "—"],
        ["Unplaced / Still in Process", "2", "Higher Studies / Research", "32", "—", "—"]
    ]
    data_table(doc, summary_headers, summary_rows, col_widths=[1.5, 0.6, 1.8, 0.6, 1.4, 1.1], compact=True)

    sub_section_header(doc, "Top Recruiting Companies & Offer Distribution", "7.2")
    body(doc, "Complete roster of all 59 corporate recruiting partners and offers made to Computer Engineering students:")

    data = load_docx_data()
    comp_table = data['7. Overview of Academic Programs.docx']['tables'][0]['rows']
    # Skip header
    comp_rows = comp_table[1:]
    data_table(doc, ["Sr. No.", "Recruiting Company", "Offers Count"], comp_rows, col_widths=[0.8, 4.7, 1.5], compact=True)

    sub_section_header(doc, "Industry Internships (2022–2026 Batch / 2025–26 Passout)", "7.3")
    body(doc, "Complete roster of all 73 corporate organizations hosting Computer Engineering students for 6-month industry internships:")

    intern_table = data['7. Overview of Academic Programs.docx']['tables'][2]['rows']
    # Row 3 is the header: ['Serial Number', 'Name of the oragnisation', 'No.of students']
    intern_rows = [r for r in intern_table[4:] if any(c.strip() for c in r)]
    data_table(doc, ["Sr. No.", "Name of the Organisation", "Students"], intern_rows, col_widths=[0.8, 4.7, 1.5], compact=True)

    sub_section_header(doc, "Research Internships at Premier Institutes", "7.4")
    body(doc, "Distribution of students engaged in cutting-edge academic and national laboratory research internships:")

    res_table = data['7. Overview of Academic Programs.docx']['tables'][3]['rows']
    res_rows = res_table[1:]
    data_table(doc, ["Premier Institution / Research Laboratory", "Students Placed"], res_rows, col_widths=[5.0, 2.0])

    doc.add_page_break()
