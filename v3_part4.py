"""
verified_builder_part4.py - Sections 14 to 19: Industry Interaction, Alumni, Activities, Student Support, SWOC, Editorial Team
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

BASE = r'C:\Users\Nirupam\Documents\magazine ce'
SCRATCH = r'C:\Users\Nirupam\.gemini\antigravity-ide\brain\d84a1836-479b-4390-be26-b801a6bee109\scratch'
MEDIA_DIR = os.path.join(SCRATCH, 'baseline_media')

def load_docx_data():
    with open(os.path.join(SCRATCH, 'docx_data.json'), 'r', encoding='utf-8') as f:
        return json.load(f)

def build_section_14(doc):
    section_header(doc, "Industry–Institute Interaction", "SECTION 14 — CORPORATE IMMERSION & CO-TEACHING")

    sub_section_header(doc, "Curriculum Industry Collaborations", "14.1")
    body(doc, "The Department actively bridges the gap between academic pedagogy and corporate enterprise through industry-curated curriculum tracks:")
    bullet(doc, "Banking Technology Track (in direct collaboration with Barclays): Prepares students for distributed transaction pipelines, high-frequency trading platforms, and financial security compliance.", bold_prefix="Barclays Collaboration:")
    bullet(doc, "Management Multidisciplinary Minor (in partnership with S.P. Jain Institute of Management and Research - SPJIMR): Provides managerial and operational economics foundations.", bold_prefix="SPJIMR Management Track:")
    bullet(doc, "Financial & Strategic Management: Industry-aligned coursework designed with Six Ladders.", bold_prefix="Six Ladders Track:")

    sub_section_header(doc, "Alumni & Industry Internship Evaluation", "14.2")
    body(doc, "As verified in the External Academic Audit (Criteria B, Page 18 & Criteria E, Page 53), alumni and industry practitioners actively evaluate student progress during full-semester industry internships:")
    bullet(doc, "In AY 2024–25, 18 distinguished alumni evaluated students during industry internships.")
    bullet(doc, "In AY 2025–26, 11 alumni and corporate evaluators conducted rigorous comprehensive technical assessments.")

    sub_section_header(doc, "Expert Talks & Co-Teaching Sessions", "14.3")
    bullet(doc, "Delivered by Dr. Maheshkumar H. Kolekar, Associate Professor, IIT Patna, on 13/10/2024 for the subject Artificial Intelligence and Machine Learning (AIML), arranged by Prof. Swapnali Kurhade.", bold_prefix="Academic Co-Teaching:")
    bullet(doc, "Delivered by Soumyadeb Bhattacharjee, Strategy and Management Consultant, on 'Exploring and Preparing for Management and Consulting Careers', arranged by Dr. Sudhir N. Dhage.", bold_prefix="Alumni Industry Talk:")

    doc.add_page_break()

def build_section_15(doc):
    section_header(doc, "Distinguished Alumni & Alumni Mentoring", "SECTION 15 — GLOBAL ALUMNI IMPACT")

    sub_section_header(doc, "Alumni Mentoring Programme (AY 2025–26)", "15.1")
    body(doc, "Verified alumni mentorship allocations from official Department PTM records (PTM Slide 15):")

    alumni_headers = ["Sr. No.", "Alumni Mentor Name", "Graduation Year", "Current Organization & Role"]
    alumni_rows = [
        ["1", "Akshay Polji", "2007", "Maersk Global"],
        ["2", "Soham Kanade", "2015", "Amazon"],
        ["3", "Rohit Naik", "2017", "Google"],
        ["4", "Poojan Turakhia", "2020", "JP Morgan Chase & Co."],
        ["5", "Yash Gaba", "2019", "Lovable Arts"],
        ["6", "Darsh Mehta", "2022", "Deutsche Bank"],
        ["7", "Tanvi Dhope", "2021", "Microsoft"]
    ]
    data_table(doc, alumni_headers, alumni_rows, col_widths=[0.8, 2.2, 1.2, 2.8])

    # Alumni graphics from baseline
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image9.png"), width=Inches(5.2), caption_text="Alumni Mentoring Network — Department of Computer Engineering")
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image10.png"), width=Inches(5.2), caption_text="Distinguished Alumni Across Global Tech Ecosystems")

    doc.add_page_break()

def build_section_16(doc):
    section_header(doc, "Departmental Activities & Student Associations", "SECTION 16 — TECHNICAL SOCIETIES & HACKATHONS")

    sub_section_header(doc, "Parent–Teacher Meeting (PTM 2025–26)", "16.1")
    ptm_pic = os.path.join(BASE, "PTM.JPG")
    add_image_safe(doc, ptm_pic, width=Inches(5.5), caption_text="Parent–Teacher Meeting (PTM) — 30th April 2026, S.P.I.T. Auditorium")
    body(doc, "The Department conducted its annual Parent–Teacher Meeting on 30th April 2026. The meeting provided parents with detailed visibility into academic progress, attendance monitoring, remedial teaching interventions, internship performance, and campus placement milestones.")

    sub_section_header(doc, "SPIT Hackathon 2026 (10th National Edition)", "16.2")
    callout_box(doc,
        "\"The 10th Edition of SPIT HACKATHON observed a massive registration of around 3,316 participants showcasing nationwide interest and competitive enthusiasm. After a rigorous shortlisting process, 78 teams were selected for the final 24-hour round at Sardar Patel Institute of Technology with a prize pool of ₹2.5 Lakhs!\"",
        title="10TH EDITION OF SPIT HACKATHON (PTM Slide 20)",
        border_color=NAVY_HEX, bg_color=LIGHT_BG
    )
    body(doc, "The hackathon featured four competitive tracks with 8 curated industry-level challenge statements: Artificial Intelligence & Machine Learning (AIML), Web Development, App Development, and Blockchain.")

    sub_section_header(doc, "Departmental Technical Events", "16.3")
    bullet(doc, "Organized on Thursday, 16th October 2025 across fourth-floor computer labs. Integrated competitive programming challenges with a traditional Housie format, testing algorithms, logic, and rapid execution.", bold_prefix="CodeHousie:")
    bullet(doc, "Frontend hackathon organized by C.S.I.-SPIT in collaboration with ACSES SPIT on 14th February 2026. Engaged 211 undergraduate students in UI/UX architecture, responsive design, and modern web frameworks.", bold_prefix="Pixel Perfect:")
    bullet(doc, "Odd Semester: Git Workshop, Code Housie, Something Decentralised. Even Semester: Pixel Perfect, ML Sprint, Code-It-Ralph, and Codefolio portfolio reviews.", bold_prefix="FACE Association Events:")

    sub_section_header(doc, "IEEE Student Branch & IIC Technical Competitions", "16.4")
    bullet(doc, "Held under the Institution's Innovation Council (IIC). Tested foundational computing logic, algorithmic reasoning, and problem-solving velocity.", bold_prefix="BrainDrain Competition 2025–26:")
    bullet(doc, "Fostered research methodologies, problem formulation, and structured proposal writing among undergraduate scholars.", bold_prefix="Research Ideation 2025–26:")

    # Event media from baseline
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image12.jpg"), width=Inches(3.2), caption_text="Departmental Technical Activities & Hackathons")
    add_image_safe(doc, os.path.join(MEDIA_DIR, "image13.jpg"), width=Inches(3.2), caption_text="Interactive Problem Solving Sessions")

    doc.add_page_break()

def build_section_17(doc):
    section_header(doc, "Student Support, Remedial Teaching & Pedagogy", "SECTION 17 — PEDAGOGICAL EXCELLENCE")

    data = load_docx_data()
    file17 = data['17. Student Support.docx']

    sub_section_header(doc, "Slow & Fast Learners Support Summary (AY 2025–26)", "17.1")
    t1 = file17['tables'][0]['rows']
    # Filter valid rows
    t1_rows = [r for r in t1[1:] if any(c.strip() for c in r)]
    data_table(doc, ["Faculty Name", "Subject", "Semester", "Intervention Type & Activities", "Hours"], t1_rows, col_widths=[1.5, 1.6, 0.7, 2.4, 0.8], compact=True)

    sub_section_header(doc, "Slow & Fast Learners Support Summary (AY 2024–25)", "17.2")
    t2 = file17['tables'][1]['rows']
    t2_rows = [r for r in t2[1:] if any(c.strip() for c in r)]
    data_table(doc, ["Faculty Name", "Subject", "Semester", "Intervention Type & Activities", "Hours"], t2_rows, col_widths=[1.5, 1.6, 0.7, 2.4, 0.8], compact=True)

    sub_section_header(doc, "Innovative Teaching–Learning Practices (AY 2025–26)", "17.3")
    t3 = file17['tables'][2]['rows']
    t3_rows = []
    for r in t3:
        if len(r) >= 6 and r[0].strip().isdigit():
            sr, name, title, subj, cls, sem = [c.strip() for c in r[:6]]
            t3_rows.append([sr, name, subj, cls, title])
    data_table(doc, ["Sr.", "Faculty Name", "Subject", "Class / Sem", "Innovative Pedagogical Intervention"], t3_rows, col_widths=[0.5, 1.6, 1.0, 1.1, 2.8], compact=True)

    sub_section_header(doc, "Innovative Teaching–Learning Practices (AY 2024–25)", "17.4")
    t4 = file17['tables'][3]['rows']
    t4_rows = []
    for r in t4:
        if len(r) >= 6 and r[0].strip().isdigit():
            sr, name, title, subj, cls, sem = [c.strip() for c in r[:6]]
            t4_rows.append([sr, name, subj, cls, title])
    data_table(doc, ["Sr.", "Faculty Name", "Subject", "Class / Sem", "Innovative Pedagogical Intervention"], t4_rows, col_widths=[0.5, 1.6, 1.0, 1.1, 2.8], compact=True)

    doc.add_page_break()

def build_section_18(doc):
    section_header(doc, "SWOC Analysis & Departmental Goals", "SECTION 18 — STRATEGIC APPRAISAL & ROADMAP")

    sub_section_header(doc, "Departmental SWOC Analysis Matrix (Audit PDF Page 78)", "18.1")

    # Strengths
    callout_box(doc,
        "• Established Academic Ecosystem: Autonomous, NEP 2020-aligned curriculum with flexible electives, experiential learning, and outcome-based education (OBE).\n• Research & Higher Studies: Active student/faculty research, IPR filing, 40 publications (15 journals + 25 conferences), and 32 students in premier higher studies.\n• Industry & Alumni Engagement: Direct corporate curriculum tracks (Barclays, SPJIMR), 11 alumni evaluators, and 98.21% placement rate.",
        title="STRENGTHS (S)",
        border_color=NAVY_HEX, bg_color=LIGHT_BG
    )

    # Weaknesses
    callout_box(doc,
        "• Sponsored Research & Grants: Need to scale sustained externally sponsored project funding, industrial consultancy, and large-scale government R&D outlays.\n• Documentation & Action Tracking: Opportunity to further systematize continuous improvement records and longitudinal mentoring logs.\n• Interdisciplinary Research: Scope to broaden collaborative cross-departmental funded initiatives.",
        title="AREAS TO STRENGTHEN / WEAKNESSES (W)",
        border_color=GOLD_HEX, bg_color=GOLD_BG_HEX
    )

    # Opportunities
    callout_box(doc,
        "• Frontier Computing Domains: Deepen curriculum and specialized laboratories in Generative AI, Cloud-Native Systems, Cybersecurity, and Quantum Computing.\n• Collaborative Research Ecosystem: Expand research partnerships with national laboratories (IIT Bombay, DRDO) and industry technology hubs.\n• Innovation & Entrepreneurship: Convert hackathon prototypes into commercial ventures, startup incubation, and patent filings.",
        title="OPPORTUNITIES (O)",
        border_color=NAVY_HEX, bg_color=LIGHT_BG
    )

    # Challenges
    callout_box(doc,
        "• Rapid Technology Lifecycles: Continuous alignment of computing curricula, software tools, and faculty proficiencies with accelerating industry paradigms.\n• Scaling Quality Mentorship: Sustaining personalized research mentoring, project guidance, and lab access with expanded student cohorts (240 intake).\n• Global Competition: Fierce competition for sponsored research grants, faculty talent acquisition, and tier-1 corporate partnerships.",
        title="CHALLENGES (C)",
        border_color=GOLD_HEX, bg_color=GOLD_BG_HEX
    )

    sub_section_header(doc, "Strategic Goals of the Department 2026–2030 (Audit PDF Page 79)", "18.2")
    bullet(doc, "Expand funded research, consultancy outlays, IPR generation, and interdisciplinary project activity.", bold_prefix="Goal 01 — Research & Sponsored Projects:")
    bullet(doc, "Deepen corporate co-teaching, sponsored labs, 6-month internships, and executive expert engagement.", bold_prefix="Goal 02 — Industry–Academia Integration:")
    bullet(doc, "Continuously strengthen learning infrastructure for AI/ML, GenAI, Quantum Computing, and Cybersecurity.", bold_prefix="Goal 03 — Emerging Technology Readiness:")
    bullet(doc, "Convert student hackathon projects and capstones into functional prototypes, research publications, and patents.", bold_prefix="Goal 04 — Innovation & Entrepreneurship:")
    bullet(doc, "Expand research internships, academic exchanges, and progression pathways to world-class universities.", bold_prefix="Goal 05 — Global & Higher-Study Engagement:")
    bullet(doc, "Use annual academic audit evidence, stakeholder feedback, and attainment data for continuous pedagogical enhancement.", bold_prefix="Goal 06 — Continuous Quality Improvement:")

    doc.add_page_break()

def build_section_19(doc):
    section_header(doc, "Editorial Board & Acknowledgments", "SECTION 19 — MEET THE EDITORIAL TEAM")

    callout_box(doc,
        "\"This Annual Magazine represents the collective intellect, artistic dedication, and academic vibrancy of the Department of Computer Engineering. We express our deepest gratitude to our institution's leadership, our faculty, staff, alumni, and our brilliant students whose achievements grace these pages.\"",
        title="EDITORIAL NOTE & CONCLUDING REFLECTIONS",
        border_color=GOLD_HEX, bg_color=GOLD_BG_HEX
    )

    sub_section_header(doc, "Editorial & Advisory Leadership", "19.1")
    ed_headers = ["Designation / Role", "Name", "Institutional Affiliation"]
    ed_rows = [
        ["Chief Patron", "Dr. Bhalchandra Chaudhari", "Principal, Sardar Patel Institute of Technology"],
        ["Editor-in-Chief", "Dr. Sudhir N. Dhage", "Professor & Head, Department of Computer Engineering"],
        ["Faculty Editorial Advisor", "Prof. Swapnali Kurhade", "Assistant Professor, Department of Computer Engineering"],
        ["Faculty Editorial Advisor", "Prof. Jyoti Ramteke", "Assistant Professor, Department of Computer Engineering"],
        ["Faculty Editorial Advisor", "Prof. Reeta Koshy", "Assistant Professor, Department of Computer Engineering"],
        ["Student Editorial Lead", "Department Student Editorial Committee", "Department of Computer Engineering, S.P.I.T."]
    ]
    data_table(doc, ed_headers, ed_rows, col_widths=[2.2, 2.5, 2.3])

    add_spacer(doc, 12)

    # Closing Footer Box
    p_close = doc.add_paragraph()
    p_close.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_styled(p_close, "BHARATIYA VIDYA BHAVAN'S\n", 'Calibri', 11, bold=True, color=GOLD_DARK)
    run_styled(p_close, "SARDAR PATEL INSTITUTE OF TECHNOLOGY\n", 'Cambria', 14, bold=True, color=NAVY)
    run_styled(p_close, "DEPARTMENT OF COMPUTER ENGINEERING\n", 'Calibri', 11, bold=True, color=NAVY)
    run_styled(p_close, "Munshi Nagar, Andheri (West), Mumbai — 400 058, Maharashtra, India\n", 'Calibri', 9.5, color=GREY)
    run_styled(p_close, "Website: https://comp.spit.ac.in/ | Autonomous Institute Affiliated to University of Mumbai\n", 'Calibri', 9, italic=True, color=GREY)
    run_styled(p_close, "© Academic Year 2025–2026. All Rights Reserved.", 'Calibri', 8.5, color=GREY)
