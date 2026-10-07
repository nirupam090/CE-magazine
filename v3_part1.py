"""
Part 1 of Version 3 Magazine:
- Cover Page
- Hierarchical Table of Contents (TOC)
- Section 01: From the Principal's Desk
- Section 02: From the HOD's Desk
- Section 03 & 04: Vision, Mission, PEOs, POs, PSOs & Bloom's Attainment Framework
"""

import os
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from v3_core import (
    BASE, IMG, NAVY, GOLD, GOLD_DARK, GREEN, WHITE, BLACK, GREY, LIGHT_GREY,
    NAVY_HEX, NAVY_DARK_HEX, GOLD_HEX, GOLD_BG_HEX, LIGHT_BG, BORDER_HEX, WHITE_HEX, ALT_ROW,
    BLUE_BG, GREEN_BG, GREEN_HEX,
    run_styled, add_spacer, section_header, sub_section_header, stat_cards,
    pull_quote, vision_box, highlight_card, data_table, gold_rule, thin_rule,
    body, body_rich, bullet, add_image_safe, caption, set_cell_shading,
    set_cell_margins, set_cell_borders, make_table_full_width
)

def build_cover_and_frontmatter(doc):
    # ── Cover Banner ──
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t)
    c = t.cell(0, 0)
    set_cell_shading(c, NAVY_DARK_HEX)
    set_cell_margins(c, top=280, bottom=240, left=240, right=240)
    set_cell_borders(c, bottom={'color': GOLD_HEX, 'sz': '24', 'style': 'single'})

    p0 = c.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_styled(p0, "BHARATIYA VIDYA BHAVAN'S", 'Calibri', 11, bold=True, color=GOLD)

    p1 = c.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(4)
    run_styled(p1, "SARDAR PATEL INSTITUTE OF TECHNOLOGY", 'Cambria', 17, bold=True, color=WHITE)

    p2 = c.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(2)
    run_styled(p2, "(An Autonomous Institute Affiliated to University of Mumbai)", 'Calibri', 9.5, italic=True, color=GOLD)

    add_spacer(doc, 14)

    # Title & Edition
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(8)
    p3.paragraph_format.space_after = Pt(2)
    run_styled(p3, "DEPARTMENT OF COMPUTER ENGINEERING", 'Cambria', 22, bold=True, color=NAVY)

    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(2)
    p4.paragraph_format.space_after = Pt(4)
    run_styled(p4, "ANNUAL MAGAZINE 2025 – 26", 'Cambria', 16, bold=True, color=GOLD_DARK)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(2)
    p_sub.paragraph_format.space_after = Pt(12)
    run_styled(p_sub, "Empowering Next-Generation Computing: Intelligence, Systems, and Ethical Innovation", 'Calibri', 11, italic=True, color=GREY)

    gold_rule(doc)

    stat_cards(doc, [
        ("4 CYCLES", "NBA ACCREDITED (2013, 2016, 2021, 2024)", "gold"),
        ("160 CR", "NEP 2020 PI-MODEL (57% CHOICE)", "navy"),
        ("98.21%", "CAMPUS PLACEMENT RATE", "green"),
        ("₹6.05 CR", "RESEARCH PROPOSALS SUBMITTED", "navy")
    ])

    pull_quote(doc,
        "Nurturing socially conscious, technologically empowered engineers equipped with sound fundamentals, "
        "interdisciplinary curiosity, and the resilience to pioneer impactful real-world solutions.",
        "Department Vision Statement")

    doc.add_page_break()

    # ═══════════════════════════════════════
    # HIERARCHICAL TABLE OF CONTENTS (TOC)
    # ═══════════════════════════════════════
    section_header(doc, "TABLE OF CONTENTS", "MAGAZINE INDEX — AY 2025–26")
    body(doc, "Structured in full alignment with the institutional department audit framework and accreditation governance.")

    toc_items = [
        ("1.0", "From the Principal's Desk", "03"),
        ("2.0", "From the HOD's Desk", "04"),
        ("3.0", "Department Vision and Mission Statements", "06"),
        ("   3.1", "Program Educational Objectives (PEOs)", "06"),
        ("4.0", "Program Outcomes (POs) & Program Specific Outcomes (PSOs)", "07"),
        ("   4.1", "National Board of Accreditation Program Outcomes (PO 1–12)", "07"),
        ("   4.2", "Program Specific Outcomes (PSO 1–2)", "08"),
        ("   4.3", "Bloom's Taxonomy Framework & Direct/Indirect Attainment Model", "08"),
        ("5.0", "About the Department — Curriculum and Syllabus Highlights", "09"),
        ("   5.1", "Redefining Engineering Education: The Liberal 'Pi-Model' Architecture", "09"),
        ("   5.2", "Specialisation Pathways, Multidisciplinary Minors & Experiential Learning", "11"),
        ("6.0", "Laboratory Facilities, Smart Classrooms & Equipment Added", "13"),
        ("   6.1", "State-of-the-Art Department Laboratories & Smart Classrooms", "13"),
        ("   6.2", "Major Equipment & Hardware Modernisation in 2025–26 (₹49.44L Upgrade)", "14"),
        ("7.0", "Overview of Academic Programs, Placements, Internships & Higher Studies", "15"),
        ("   7.1", "Academic Cohorts (UG B.Tech, PG M.Tech, Ph.D.) & Statistical Enrollment", "15"),
        ("   7.2", "Campus Placements 2025–26 (98.21% Placed, Salary Tiers & 60 Recruiters)", "16"),
        ("   7.3", "Industry Internships (Batch 2022–2026) & Global Higher Studies Progression", "18"),
        ("8.0", "Faculty and Staff Information", "19"),
        ("   8.1", "Faculty Profile & Doctoral Strength (30 Faculty, 12 Ph.D., 6 Pursuing)", "19"),
        ("   8.2", "Technical, Laboratory & Administrative Supporting Staff", "21"),
        ("9.0", "Additional Responsibilities Entrusted to the Faculty", "22"),
        ("10.0", "Departmental Governance: Advisory Board, Board of Studies & DAPC", "23"),
        ("11.0", "Consultancy, Grants, Patents and Research Publications", "24"),
        ("   11.1", "Faculty Citation Metrics (1,077 Citations, h-index & i10-index Roster)", "24"),
        ("   11.2", "Research Grant Proposals (10 Proposals worth ₹6.05 Cr) & Consultancy (₹10.2L)", "25"),
        ("   11.3", "Intellectual Property & Patents Filed (Sound2Sketch & AI Monitoring)", "26"),
        ("   11.4", "International Journal & Conference Publications (Scopus, SCIE, IEEE Xplore)", "27"),
        ("12.0", "Faculty and Students Achievements", "28"),
        ("   12.1", "Faculty Honors, NPTEL Star Certifications & Best Paper Awards", "28"),
        ("   12.2", "Student Hackathons (EY Techathon ₹3L Winner, Barclays, Google Cloud GenAI)", "29"),
        ("   12.3", "GATE DA AIR-97, Bug Bounty Halls of Fame (Ethereum & Zoho) & Academic Honors", "31"),
        ("13.0", "List of Undergraduate Projects, M.Tech Dissertations & Ph.D. Research", "33"),
        ("   13.1", "B.Tech Major Capstone Projects 2025–26 (57 Groups Title, Guide & Domain)", "33"),
        ("   13.2", "M.Tech Dissertation Research Projects (Agentic RAG, Zero Trust, Smart Grid)", "36"),
        ("   13.3", "Capstone Project Domain Distribution & NTRO Defense Research Spotlight", "37"),
        ("14.0", "Industry–Institute Interaction & Professional Collaborations", "38"),
        ("   14.1", "Industry MDMs (Barclays Banking Tech, SPJIMR Management, Vizuara, NVIDIA DLI)", "38"),
        ("   14.2", "Co-Teaching Initiatives (IIT Patna) & Corporate Expert Lectures", "39"),
        ("   14.3", "Faculty Professional Development (8,600+ Hours STTP/FDP) & Active MOUs", "40"),
        ("15.0", "Distinguished Alumni & Student Entrepreneurs", "41"),
        ("   15.1", "Alumni Mentoring Program & Project Evaluation Panel (18 in 24-25, 11 in 25-26)", "41"),
        ("   15.2", "Distinguished Alumni Roster Across Global Tech Giants & Enterprises", "42"),
        ("   15.3", "Alumni Entrepreneurs & Startup Ventures", "43"),
        ("16.0", "Departmental Activities by Faculty, Students and Associations", "44"),
        ("   16.1", "Parent-Teacher Meeting (PTM 2026) & Stakeholder Engagement", "44"),
        ("   16.2", "CSI SPIT Chapter (SPIT Hackathon 10th Ed, CodeHousie, Pixel Perfect)", "45"),
        ("   16.3", "IEEE Computer Society (Brain Drain, Research Ideation, Quantum Conclave)", "47"),
        ("   16.4", "FACE Initiatives (Donation Drives 1 & 2, FaceOff, Codefolio, LLC Exhibits)", "49"),
        ("17.0", "Student Support Systems, Remedial Teaching & Pedagogical Innovations", "50"),
        ("   17.1", "Structured Support for Slow Learners & Challenges for Fast Learners (~54 hrs)", "50"),
        ("   17.2", "Innovative Teaching–Learning Practices by Faculty Roster (Card Tricks, Flipped)", "51"),
        ("   17.3", "360-Degree Mentorship, Bloom's Taxonomy & Open House Transparent Evaluation", "53"),
        ("18.0", "Departmental Strengths, Weaknesses, Opportunities & Challenges (SWOC)", "54"),
        ("   18.1", "Comprehensive SWOC Analysis Matrix (4 Quadrants)", "54"),
        ("   18.2", "Core Institutional Pillars & Distinctive Department Strengths", "55"),
        ("   18.3", "Strategic Roadmap 2026–2030 (6 Strategic Horizon Goals)", "56"),
        ("19.0", "Meet the Team — Editorial Board & Acknowledgments", "57")
    ]

    t_toc = doc.add_table(rows=1 + len(toc_items), cols=3)
    t_toc.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t_toc)

    # Header
    for j, h in enumerate(["SR. NO.", "SECTION / SUB-SECTION TITLE", "PAGE"]):
        cell = t_toc.cell(0, j)
        set_cell_shading(cell, NAVY_HEX)
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        set_cell_borders(cell, bottom={'color': GOLD_HEX, 'sz': '12', 'style': 'single'})
        p = cell.paragraphs[0]
        run_styled(p, h, 'Calibri', 9, bold=True, color=WHITE)

    for i, (num, title_txt, pg) in enumerate(toc_items):
        cell_num = t_toc.cell(i + 1, 0)
        cell_title = t_toc.cell(i + 1, 1)
        cell_pg = t_toc.cell(i + 1, 2)
        is_sub = num.startswith("   ")
        bg = ALT_ROW if i % 2 == 1 else WHITE_HEX

        for cell in [cell_num, cell_title, cell_pg]:
            set_cell_shading(cell, bg)
            set_cell_margins(cell, top=35, bottom=35, left=80, right=80)
            set_cell_borders(cell, bottom={'color': BORDER_HEX, 'sz': '1', 'style': 'single'})

        p_n = cell_num.paragraphs[0]
        run_styled(p_n, num.strip(), 'Calibri', 8.5, bold=(not is_sub), color=NAVY if not is_sub else GREY)

        p_t = cell_title.paragraphs[0]
        if is_sub:
            p_t.paragraph_format.left_indent = Inches(0.2)
            run_styled(p_t, title_txt, 'Calibri', 8.5, italic=True, color=BLACK)
        else:
            run_styled(p_t, title_txt, 'Calibri', 9, bold=True, color=NAVY)

        p_p = cell_pg.paragraphs[0]
        p_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_styled(p_p, pg, 'Calibri', 8.5, bold=(not is_sub), color=GOLD_DARK if not is_sub else GREY)

    add_spacer(doc, 10)
    doc.add_page_break()

def build_section_01(doc):
    # ═══════════════════════════════════════
    # SECTION 01 — FROM THE PRINCIPAL'S DESK
    # ═══════════════════════════════════════
    section_header(doc, "From the Principal's Desk", "SECTION 01 — INSTITUTIONAL LEADERSHIP")

    # Layout with Principal Photo
    t_box = doc.add_table(rows=1, cols=2)
    t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    make_table_full_width(t_box)

    c_img = t_box.cell(0, 0)
    c_img.width = Inches(2.2)
    set_cell_shading(c_img, LIGHT_BG)
    set_cell_margins(c_img, top=100, bottom=100, left=100, right=100)
    set_cell_borders(c_img,
        top={'color': NAVY_HEX, 'sz': '12', 'style': 'single'},
        left={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})

    p_img = c_img.paragraphs[0]
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    princ_photo = os.path.join(IMG, 'principal_photo.png')
    if os.path.exists(princ_photo):
        r = p_img.add_run()
        r.add_picture(princ_photo, width=Inches(1.8))
    p_lbl = c_img.add_paragraph()
    p_lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_lbl.paragraph_format.space_before = Pt(4)
    run_styled(p_lbl, "Dr. Bhalchandra Chaudhari", 'Cambria', 10, bold=True, color=NAVY)
    p_sub = c_img.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(1)
    run_styled(p_sub, "Principal\nSardar Patel Institute of Technology", 'Calibri', 8, italic=True, color=GREY)

    c_txt = t_box.cell(0, 1)
    set_cell_shading(c_txt, WHITE_HEX)
    set_cell_margins(c_txt, top=100, bottom=100, left=140, right=140)
    set_cell_borders(c_txt,
        top={'color': GOLD_HEX, 'sz': '12', 'style': 'single'},
        right={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        bottom={'color': BORDER_HEX, 'sz': '2', 'style': 'single'},
        left={'color': BORDER_HEX, 'sz': '2', 'style': 'single'})

    p_greet = c_txt.paragraphs[0]
    p_greet.paragraph_format.space_after = Pt(4)
    run_styled(p_greet, "Dear Students, Faculty Colleagues, Esteemed Alumni, and Well-Wishers,", 'Calibri', 10, bold=True, color=NAVY)

    p_body1 = c_txt.add_paragraph()
    p_body1.paragraph_format.space_after = Pt(4)
    run_styled(p_body1,
        "It gives me immense pride and profound joy to present the Annual Magazine of the Department of Computer Engineering for the Academic Year 2025–26. "
        "At Sardar Patel Institute of Technology, we have consistently envisioned higher technical education not merely as the transmission of engineering curricula, "
        "but as an organic, holistic development of character, intellectual rigor, ethical grounding, and human empathy.", 'Calibri', 9.5)

    p_body2 = c_txt.add_paragraph()
    p_body2.paragraph_format.space_after = Pt(4)
    run_styled(p_body2,
        "As an Empowered Autonomous Institution, S.P.I.T. has embraced the transformative principles of the National Education Policy (NEP 2020) through our pioneering "
        "Liberal 'Pi-Model' of Engineering Education. This architecture balances rigorous computing fundamentals with 57% choice-based flexibility, multidisciplinary minor "
        "tracks in banking technology, management, and industrial IoT, and liberal learning in ethics, yoga, and literature. "
        "Our students do not just write code; they design solutions that elevate community life and address the pressing imperatives of our nation.", 'Calibri', 9.5)

    add_spacer(doc, 6)

    body(doc,
        "The achievements chronicled across these pages bear eloquent testimony to the vibrant intellectual spirit of the Department of Computer Engineering. "
        "In Academic Year 2025–26, the department has achieved an extraordinary 98.21% placement rate, with graduates securing premier positions across top multinational "
        "corporations and global technology titans with compensation packages reaching ₹44.0 LPA. Simultaneously, our student innovators have distinguished themselves "
        "nationally by capturing top honors at the EY Techathon 6.0 (₹3 Lakhs First Prize), the Barclays Hackathon, Google Cloud GenAI Exchange, and global Bug Bounty Halls of Fame "
        "at Ethereum and Zoho, alongside stellar national academic feats including All India Rank 97 in GATE Data Science and Artificial Intelligence.")

    body(doc,
        "On the research and academic infrastructure frontier, the department has taken significant strides forward. The state-of-the-art laboratory modernisation "
        "with 80 HP ProDesk 400 G7 workstations backed by ₹49.44 Lakhs of institutional investment, combined with the submission of 10 competitive research grant proposals "
        "worth ₹6.05 Crores to agencies such as DRDO, SERB, and AICTE, highlights our unwavering commitment to world-class research. "
        "Furthermore, our faculty members have contributed over 40 research papers in Scopus and SCI-indexed venues, earning more than 1,077 citations and driving impactful patents.")

    pull_quote(doc,
        "Education at S.P.I.T. is not merely about assimilating technical knowledge; it is about cultivating character, ethical discipline, "
        "and the courageous imagination to build transformative solutions for humanity.",
        "Dr. Bhalchandra Chaudhari, Principal")

    body(doc,
        "I commend Dr. Sudhir N. Dhage, Head of the Department, along with the faculty members, technical staff, student editorial leads, and the entire departmental family "
        "for their dedicated stewardship in chronicling this remarkable year of excellence. May this publication serve as an enduring beacon of inspiration, reminding each reader "
        "of the heights that can be scaled when vision, diligence, and integrity unite.")

    doc.add_page_break()

def build_section_02(doc):
    # ═══════════════════════════════════════
    # SECTION 02 — FROM THE HOD'S DESK
    # ═══════════════════════════════════════
    section_header(doc, "From the HOD's Desk", "SECTION 02 — DEPARTMENTAL LEADERSHIP")

    pull_quote(doc,
        "Our mission is to transform undergraduate computing from passive curriculum completion into active creation—where every student "
        "engineers solutions that matter to industry and society.",
        "Dr. Sudhir N. Dhage, Professor & Head, Department of Computer Engineering")

    body(doc, "Dear Students, Faculty Colleagues, Esteemed Alumni, and Respected Partners,", bold=True, color=NAVY)

    body(doc,
        "It is my distinct privilege to welcome you to the 2025–26 edition of our Annual Magazine. As we reflect on another defining year in the life of the "
        "Department of Computer Engineering at Bharatiya Vidya Bhavan's Sardar Patel Institute of Technology, I am filled with immense pride at the exceptional "
        "resilience, creativity, and intellectual mastery demonstrated across every facet of our departmental ecosystem.")

    body(doc,
        "The computer engineering landscape is navigating a watershed epoch marked by the pervasive convergence of Artificial Intelligence, Generative Foundation Models, "
        "Quantum Machine Learning, Autonomous Distributed Systems, and Zero-Trust Cybersecurity. In response, our department has continued to spearhead educational reform. "
        "Through our NEP 2020 Liberal 'Pi-Model' curriculum, we have established 57 choice-based academic pathways across 160 credits, allowing our students to integrate "
        "deep computing specialisations with industry-embedded Multidisciplinary Minors (such as Banking Technology with Barclays and Management with SPJIMR) and liberal arts. "
        "This progressive structure ensures our graduates are neither narrow coders nor disconnected theorists, but well-rounded innovators capable of addressing multidimensional societal challenges.")

    stat_cards(doc, [
        ("98.21%", "PLACEMENT SUCCESS (110 / 112)", "green"),
        ("₹49.44L", "LAB UPGRADE (80 HP PRODESKS)", "navy"),
        ("₹6.05 CR", "GRANT PROPOSALS (10 SUBMISSIONS)", "gold"),
        ("1,077+", "FACULTY RESEARCH CITATIONS", "navy")
    ])

    body(doc,
        "The Academic Year 2025–26 has witnessed groundbreaking milestones across research, placements, infrastructure, and student innovation:")

    bullet(doc, "98.21% Placement Rate with 110 placed out of 112 eligible students across top-tier firms including Barclays (21 offers), Nomura, BlackRock, Morgan Stanley, JP Morgan Chase, and PhonePe, with average compensation reaching ₹14.85 LPA and peak domestic offers of ₹44.0 LPA.", bold_prefix="Unprecedented Placements: ")
    bullet(doc, "Hardware Modernisation worth ₹49,44,388, deploying 80 high-performance HP ProDesk 400 G7 workstations with 10th Gen Core i5 processors, NVMe SSDs, and IPS displays across Computer Labs 1 & 2, complemented by prior infrastructure additions totaling ₹23,12,800.", bold_prefix="State-of-the-Art Infrastructure: ")
    bullet(doc, "Submitting 10 major research grant proposals worth ₹6.05 Crores to premier national funding agencies (DRDO, SERB CRG, AICTE RPS, and RGSTC), alongside Dr. Anant Nimkar's ₹10.2 Lakhs industrial consultancy project.", bold_prefix="Research Grant Surge: ")
    bullet(doc, "National Hackathon Crowns including the EY Techathon 6.0 National Championship (₹3 Lakhs Cash Award), Barclays Hackathon First Prize, Google Cloud GenAI Exchange 3rd Runner Up, GATE DA All India Rank 97 by Atharva Jagtap, ICPC Asia Amritapuri Regionalist representation, and global Bug Bounty Halls of Fame at Ethereum and Zoho by Madhav Shah.", bold_prefix="Student Honors: ")
    bullet(doc, "40 publications (15 International Journals, 25 International Conferences) in Scopus and SCI/SCIE indexed venues, accumulating over 1,077 citations, with 12 faculty holding Ph.D. degrees, 6 actively pursuing doctorates, and faculty completing over 8,600 hours of professional FDP/STTP training.", bold_prefix="Scholarly Rigor: ")

    body(doc,
        "Beyond individual academic triumphs, our student bodies—the Computer Society of India (CSI), IEEE Computer Society, and the Forum of Aspiring Computer Engineers (FACE)—have "
        "energised campus life. From the 10th National SPIT Hackathon attracting 3,316 participants across India to code contests like CodeHousie and Pixel Perfect, and compassionate "
        "social donation drives, our students continually demonstrate that technical prowess shines brightest when wedded to collaborative leadership and societal empathy.")

    body(doc,
        "I extend my heartfelt gratitude to our visionary Principal, Dr. Bhalchandra Chaudhari, the management of Bharatiya Vidya Bhavan, our dedicated faculty members, "
        "hardworking technical staff, enthusiastic students, supportive parents, and distinguished alumni. I invite you to explore this edition and celebrate the enduring excellence "
        "of the Computer Engineering Department.")

    doc.add_page_break()

def build_section_03_and_04(doc):
    # ═══════════════════════════════════════
    # SECTION 03 & 04 — VISION, MISSION, PEOs, POs, PSOs & BLOOM'S ATTAINMENT
    # ═══════════════════════════════════════
    section_header(doc, "Vision, Mission, PEOs, POs & PSOs", "SECTIONS 03 & 04 — ACADEMIC FOUNDATIONS")

    sub_section_header(doc, "Department Vision and Mission Statements", "3.0")

    vision_box(doc, "DEPARTMENT VISION",
        "To be a centre of excellence in Computer Engineering education and research, fostering technologically competent, "
        "ethically sound, and socially responsible professionals who innovate and lead in a dynamic global environment.")

    body(doc, "The department translates its vision into daily practice through four fundamental mission pillars:", bold=True, color=NAVY)

    mission_pillars = [
        ("M1: Quality Technical Education", "Provide state-of-the-art undergraduate and postgraduate education grounded in solid computing fundamentals, mathematical rigor, and hands-on laboratory experiences."),
        ("M2: Research & Innovation Ecosystem", "Foster an active research culture among faculty and students, encouraging interdisciplinary inquiry, publications, intellectual property generation, and translational prototypes."),
        ("M3: Industry-Academia Symbiosis", "Cultivate sustained collaborations with industry leaders through multidisciplinary minors, co-teaching, internships, and consultancies to ensure contemporary curriculum alignment."),
        ("M4: Ethical Leadership & Societal Impact", "Instill professional ethics, environmental consciousness, collaborative spirit, and lifelong learning attitudes to develop leaders who deploy computing for social good.")
    ]

    for title, desc in mission_pillars:
        highlight_card(doc, title, desc, accent_hex=GOLD_HEX, bg_hex=GOLD_BG_HEX)

    sub_section_header(doc, "Program Educational Objectives (PEOs)", "3.1")
    body(doc, "The Program Educational Objectives represent broad career and professional accomplishments that graduates are expected to achieve within 3 to 5 years of graduation:")

    peos = [
        ("PEO 1: Technical Mastery & Problem Solving", "Graduates will apply foundational and advanced knowledge of Computer Engineering to analyze, design, and implement robust, scalable, and secure computational solutions for real-world challenges."),
        ("PEO 2: Continuous Learning & Research Orientation", "Graduates will adapt to rapidly evolving technologies through independent learning, professional certifications, higher studies, and applied research in emerging computing domains."),
        ("PEO 3: Professional Ethics & Societal Responsibility", "Graduates will practice engineering with high ethical standards, environmental sensitivity, and legal compliance, ensuring that technological solutions benefit society equitably."),
        ("PEO 4: Leadership, Collaboration & Entrepreneurship", "Graduates will demonstrate effective communication, multidisciplinary teamwork, project management capabilities, and entrepreneurial initiative in corporate and startup environments.")
    ]

    for title, desc in peos:
        bullet(doc, desc, bold_prefix=f"{title}: ")

    thin_rule(doc)

    sub_section_header(doc, "Program Outcomes (POs) — National Board of Accreditation", "4.1")
    body(doc, "In accordance with the Washington Accord and NBA guidelines, every Computer Engineering graduate achieves 12 Program Outcomes upon graduation:")

    po_data = [
        ("PO 1", "Engineering Knowledge", "Apply mathematics, science, engineering fundamentals, and computer engineering specialisation to solve complex engineering problems."),
        ("PO 2", "Problem Analysis", "Identify, formulate, review research literature, and analyze complex computer engineering problems reaching substantiated conclusions."),
        ("PO 3", "Design/Development", "Design solutions for complex computational problems and design system components or processes that meet specified needs with appropriate consideration for safety and society."),
        ("PO 4", "Conduct Investigations", "Use research-based knowledge and research methods including design of experiments, analysis and interpretation of data, and synthesis of information to provide valid conclusions."),
        ("PO 5", "Modern Tool Usage", "Create, select, and apply appropriate techniques, resources, and modern engineering and IT tools including prediction and modeling to complex engineering activities."),
        ("PO 6", "The Engineer and Society", "Apply reasoning informed by contextual knowledge to assess societal, health, safety, legal, and cultural issues and consequent responsibilities relevant to professional engineering practice."),
        ("PO 7", "Environment & Sustainability", "Understand the impact of professional engineering solutions in societal and environmental contexts, and demonstrate knowledge of and need for sustainable development."),
        ("PO 8", "Ethics", "Apply ethical principles and commit to professional ethics and responsibilities and norms of engineering practice."),
        ("PO 9", "Individual and Teamwork", "Function effectively as an individual, and as a member or leader in diverse teams, and in multidisciplinary settings."),
        ("PO 10", "Communication", "Communicate effectively on complex engineering activities with the engineering community and society at large, such as being able to comprehend and write effective reports and documentation."),
        ("PO 11", "Project Management & Finance", "Demonstrate knowledge and understanding of engineering and management principles and apply these to one's own work, as a member and leader in a team, to manage projects in multidisciplinary environments."),
        ("PO 12", "Life-long Learning", "Recognize the need for, and have the preparation and ability to engage in independent and life-long learning in the broadest context of technological change.")
    ]

    data_table(doc, ["PO", "Graduate Attribute", "Description"], po_data, col_widths=[0.8, 1.8, 4.4], compact=True)

    sub_section_header(doc, "Program Specific Outcomes (PSOs)", "4.2")
    body(doc, "Complementing the general engineering outcomes, the Department of Computer Engineering specifies two distinctive Program Specific Outcomes tailored to contemporary industry demands:")

    highlight_card(doc, "PSO 1: Systems & Software Architecture",
        "Graduates will possess the ability to architect, analyze, design, and deploy secure, high-performance computing systems, "
        "scalable cloud-native infrastructures, distributed software applications, and microservice frameworks adhering to industry design patterns.",
        accent_hex=NAVY_HEX, bg_hex=BLUE_BG)

    highlight_card(doc, "PSO 2: Data Science & Intelligent Systems",
        "Graduates will possess the competence to develop, evaluate, and operationalize intelligent algorithms leveraging Machine Learning, "
        "Deep Learning, Generative AI, Big Data analytics, Computer Vision, and Natural Language Processing to solve complex data-driven industry challenges.",
        accent_hex=GREEN_HEX, bg_hex=GREEN_BG)

    sub_section_header(doc, "Bloom's Taxonomy Framework & Direct/Indirect Attainment Model", "4.3")
    body(doc,
        "The Department follows an outcome-based education (OBE) model governed by a rigorous continuous evaluation and attainment architecture. "
        "Course Outcomes (COs) are mapped to POs and PSOs with designated correlation levels (1: Low, 2: Medium, 3: High). "
        "Student learning is evaluated across Bloom's Cognitive Hierarchy from L1 (Remembering) to L6 (Creating).")

    bullet(doc, "Direct Attainment (80% Weightage): Evaluated via Continuous Assessment (In-Semester Tests, Laboratory Journals, Mini Projects, Assignments) and End-Semester Examinations (ESE).", bold_prefix="Assessment Mix: ")
    bullet(doc, "Indirect Attainment (20% Weightage): Evaluated through Course Exit Surveys, Graduate Exit Surveys, Alumni Surveys, and Employer Feedback.", bold_prefix="Stakeholder Input: ")
    bullet(doc, "Two-Set ESE Question Paper Policy: Question papers are submitted in two parallel sets with complete Bloom's taxonomy mapping and marking schemes. A departmental moderation committee ensures rigorous cognitive distribution.", bold_prefix="Paper Setting & Moderation: ")
    bullet(doc, "Open House Answer Script Evaluation: Following semester exams, an Open House session allows students to personally review evaluated scripts, discuss scoring with faculty mentors, and verify transparent marking before grades are finalized.", bold_prefix="Evaluation Transparency: ")
    bullet(doc, "Statistical Mark Analysis: Performance across core courses is mapped against Gaussian score distributions (e.g., Operating Systems semester average of SA=73) to identify outliers and drive continuous curricular improvements.", bold_prefix="Statistical Benchmarking: ")

    doc.add_page_break()
