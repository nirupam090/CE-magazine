"""
verified_builder_part3.py - Sections 8 to 13: Faculty Roster, Governance, Research Publications, Achievements, Major Projects
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

def load_docx_data():
    with open(os.path.join(SCRATCH, 'docx_data.json'), 'r', encoding='utf-8') as f:
        return json.load(f)

def build_section_08(doc):
    section_header(doc, "Faculty and Staff Information", "SECTION 08 — ACADEMIC FACULTY & CADRE")

    # Faculty Group Picture
    fac_pic = os.path.join(BASE, "Faclty group pic.JPG")
    add_image_safe(doc, fac_pic, width=Inches(6.2), caption_text="Faculty Group — Department of Computer Engineering, Academic Year 2025–26")
    add_spacer(doc, 4)

    sub_section_header(doc, "Teaching Faculty Roster (Audited Roster 2025–26)", "8.1")
    body(doc, "The Department of Computer Engineering comprises 30 full-time faculty members, structured across academic cadres as verified in the External Academic Audit (Audit PDF Page 3):")

    fac_headers = ["Sr. No.", "Name of Faculty Member", "Designation"]
    fac_rows = [
        ["1", "Dr. Sudhir Dhage", "Professor & Head of Department"],
        ["2", "Dr. Surekha Dholay", "Associate Professor"],
        ["3", "Dr. Anant Nimkar", "Associate Professor"],
        ["4", "Dr. Prasenjit Bhavathankar", "Associate Professor"],
        ["5", "Prof. Anand Godbole", "Associate Professor"],
        ["6", "Dr. Nataasha Raul", "Assistant Professor"],
        ["7", "Dr. Pramod Bide", "Assistant Professor"],
        ["8", "Dr. Kailas Devadkar", "Assistant Professor"],
        ["9", "Dr. Disha Sail", "Assistant Professor"],
        ["10", "Dr. Jignesh Sisodia", "Assistant Professor"],
        ["11", "Prof. Reeta Koshy", "Assistant Professor"],
        ["12", "Prof. Jyoti Ramteke", "Assistant Professor"],
        ["13", "Prof. Sunil Ghane", "Assistant Professor"],
        ["14", "Prof. Swapnali Kurhade", "Assistant Professor"],
        ["15", "Prof. Abhijeet Salunke", "Assistant Professor"],
        ["16", "Dr. Anuj Tawari", "Assistant Professor"],
        ["17", "Ms. Soni Bhambar", "Assistant Professor"],
        ["18", "Prof. Govind Gaundalkar", "Assistant Professor"],
        ["19", "Ms. Yogita Patil", "Assistant Professor"],
        ["20", "Prof. Azman Qureshi", "Assistant Professor"],
        ["21", "Prof. Jotsna Bhagat", "Assistant Professor"],
        ["22", "Dr. Sanjuktarani Jena", "Assistant Professor"],
        ["23", "Mr. Sudeep Mali", "Assistant Professor"],
        ["24", "Prof. Isha Sawalkar", "Assistant Professor"],
        ["25", "Mr. Ameya Nile", "Assistant Professor"],
        ["26", "Mr. Amey Morye", "Assistant Professor"],
        ["27", "Prof. Vishwas Badhe", "Assistant Professor"],
        ["28", "Ms. Aishwarya Nalawade", "Assistant Professor"],
        ["29", "Ms. Falak Malik", "Assistant Professor"],
        ["30", "Mr. Anas Ansari", "Assistant Professor"]
    ]
    data_table(doc, fac_headers, fac_rows, col_widths=[0.8, 3.8, 2.4], compact=True)

    sub_section_header(doc, "Technical & Laboratory Support Staff", "8.2")
    tech_headers = ["Sr. No.", "Name of Staff Member", "Designation"]
    tech_rows = [
        ["1", "Mr. Rajendra Nalawade", "Technical Assistant"],
        ["2", "Mr. Nandankumar Kinare", "Technical Assistant"],
        ["3", "Mrs. Rupali Bhosale", "Technical Assistant"]
    ]
    data_table(doc, tech_headers, tech_rows, col_widths=[0.8, 3.8, 2.4])

    sub_section_header(doc, "Supporting Staff", "8.3")
    supp_headers = ["Sr. No.", "Name of Staff Member", "Designation"]
    supp_rows = [
        ["1", "Mr. Mangesh Birje", "Supporting Staff"],
        ["2", "Mr. Vishwanath Kamblekar", "Supporting Staff"],
        ["3", "Mr. Prathamesh Bhonde", "Supporting Staff"]
    ]
    data_table(doc, supp_headers, supp_rows, col_widths=[0.8, 3.8, 2.4])

    doc.add_page_break()

def build_section_09_and_10(doc):
    section_header(doc, "Faculty Responsibilities & Departmental Governance", "SECTIONS 09 & 10 — INSTITUTIONAL GOVERNANCE")

    sub_section_header(doc, "Additional Responsibilities Assigned to Faculty Members", "9.1")
    body(doc, "Faculty members lead core institutional and departmental administrative portfolios:")

    resp_headers = ["Faculty Member", "Administrative & Portfolio Responsibilities"]
    resp_rows = [
        ["Dr. Sudhir N. Dhage", "Head of Department, Chairman Board of Studies (BoS), Chairman DAPC"],
        ["Dr. Prasenjit Bhavathankar", "NAAC Coordinator, NBA Documentation Lead"],
        ["Dr. Surekha Dholay", "Technical Program Committee Reviewer, Research Coordinator"],
        ["Dr. Anant V. Nimkar", "NBA Accreditation Coordinator"],
        ["Dr. Kailas Devadkar", "Patent & IPR Cell Coordinator"],
        ["Prof. Anand Godbole", "NPTEL Coordinator, Course Coordinator (DSGT)"],
        ["Prof. Reeta Koshy", "DAPC Member, Result Analysis Coordinator"],
        ["Prof. Sunil Ghane", "Project Coordinator (BE), IQAC Committee Member"],
        ["Prof. Swapnali Kurhade", "Placement Coordinator (CE), Slow-Fast Learner Coordinator"],
        ["Prof. Jyoti Ramteke", "Co-Teaching Coordinator, Industry Interaction Lead"],
        ["Prof. Govind Gaundalkar", "Examination Cell Coordinator (CE), LMS Moodle Admin"],
        ["Prof. Yogita Patil", "Liberal Learning Course (LLC) Coordinator"],
        ["Prof. Abhijeet Salunke", "Laboratory Coordinator, Assessment Tools Lead"]
    ]
    data_table(doc, resp_headers, resp_rows, col_widths=[2.4, 4.6])

    sub_section_header(doc, "Board of Studies (BoS) & Curriculum Governance", "10.1")
    body(doc, "The Board of Studies (BoS) in Computer Engineering is chaired by Dr. Sudhir N. Dhage and comprises eminent academicians from premier institutions (IITs, NITs), distinguished industry technology leaders, and departmental senior faculty. The BoS meets bi-annually to review syllabus architectures, validate course outcomes (COs), and approve industry tracks such as Banking Technology (Barclays) and Management Minors (SPJIMR).")

    sub_section_header(doc, "Department Academic Program Committee (DAPC)", "10.2")
    body(doc, "As detailed in the External Audit Report (Criteria E, Page 52), the DAPC is responsible for academic planning, approving lesson plans within 2–3 weeks of semester commencement, and establishing Standard Assessment (SA) weightings for Mid-Semester Examinations (MSE), Continuous Internal Assessments (ISE), and End-Semester Examinations (ESE).")

    sub_section_header(doc, "Academic Monitoring Committee & 360-Degree Feedback System", "10.3")
    bullet(doc, "Academic Monitoring Committee: Comprises three members and one coordinator. Conducts random inspections of theory classes and laboratories, monitors punctuality, timetable adherence, lab software readiness, and submits fortnightly consolidated reports to the HoD.", bold_prefix="Monitoring Mechanism:")
    bullet(doc, "360-Degree Feedback: Formal course exit feedback at the end of each semester, mid-semester student evaluations, Class Representative (CR) grievance redressal sessions, and post-examination course attainment reviews.", bold_prefix="Feedback Channels:")

    doc.add_page_break()

def build_section_11(doc):
    section_header(doc, "Consultancy, Grants, Patents & Research Publications", "SECTION 11 — SCHOLARLY INQUIRY & RESEARCH")

    data = load_docx_data()
    file11 = data['11. Faculty_Students Consultancy, Grant, Patent and Research Publications.docx']

    sub_section_header(doc, "Research Output Summary (AY 2025–26)", "11.1")
    stat_cards(doc, [
        ("01", "Patent Filed"),
        ("15", "Journal Papers"),
        ("25", "Conference Papers"),
        ("09", "Scopus Indexed"),
        ("03", "SCIE Indexed")
    ])

    body(doc, "Summary of Research Output from Official Audit Data (Table 1 of Source):")
    out_headers = ["Sr. No.", "Research Output Category", "Audited Count"]
    out_rows = [
        ["1", "Consultancy", "—"],
        ["2", "Grant", "—"],
        ["3", "Patent", "01"],
        ["4", "International Journal", "15"],
        ["5", "Book", "—"],
        ["6", "Book Chapter", "—"],
        ["7", "International Conference", "25"]
    ]
    data_table(doc, out_headers, out_rows, col_widths=[1.0, 4.5, 1.5])

    sub_section_header(doc, "Faculty Citations Roster", "11.2")
    body(doc, "Official Google Scholar citation counts for departmental faculty members as compiled in the annual report:")

    cit_table = file11['tables'][0]['rows']
    # Skip multi-level header rows (rows 0 and 1)
    cit_rows = []
    for r in cit_table[2:]:
        if len(r) >= 6:
            sr, name, c2025, c2026, h_idx, i10 = r[:6]
            if name.strip():
                cit_rows.append([sr.strip(), name.strip(), c2025.strip() or "—", c2026.strip() or "—", h_idx.strip() or "—", i10.strip() or "—"])

    data_table(doc, ["Sr.", "Name of Faculty", "Citations (2025)", "Citations (2026)", "h-Index", "i10-Index"], cit_rows, col_widths=[0.6, 2.6, 1.1, 1.1, 0.8, 0.8], compact=True)

    sub_section_header(doc, "Patent Details", "11.3")
    callout_box(doc,
        "Patent Title: AI-Powered Interactive Voice-Enabled Drawing Tool\nStatus: Filed & Published\nInventors: Faculty & Student Researchers, Department of Computer Engineering, S.P.I.T.",
        title="PATENT GRANTED / FILED IN AY 2025–26",
        border_color=GOLD_HEX, bg_color=GOLD_BG_HEX
    )

    sub_section_header(doc, "International Journal Publications (Scopus / SCIE / ESCI)", "11.4")
    body(doc, "Complete list of 15 international journal publications (Table 2 of source document):")

    j_table = file11['tables'][2]['rows']
    # Header is row 0; Total is last row
    j_rows = []
    for r in j_table[1:-1]:
        if len(r) >= 10:
            sr, title, author, pub, yr, issn, link, scopus, scie, esci = r[:10]
            if title.strip():
                idx_str = []
                if scopus.strip().upper() == 'YES': idx_str.append('Scopus')
                if scie.strip().upper() == 'YES': idx_str.append('SCIE')
                if esci.strip().upper() == 'YES': idx_str.append('ESCI')
                idx_disp = ", ".join(idx_str) if idx_str else "Peer-Reviewed"
                j_rows.append([sr.strip(), title.strip(), author.strip(), pub.strip(), yr.strip(), idx_disp])

    data_table(doc, ["Sr.", "Title of Paper", "Author(s)", "Journal / Publication", "Year", "Indexing"], j_rows, col_widths=[0.5, 2.6, 1.4, 1.5, 0.5, 0.7], compact=True)

    sub_section_header(doc, "International Conference Publications", "11.5")
    body(doc, "Complete roster of 25 conference papers presented and published in IEEE and Springer international conferences:")

    conf_papers = [
        ["1", "Physics-Based Augmented Reality Framework for Interactive Home Furnishing", "Dr. Anant Nimkar", "2026 IEEE IATMSI, Gwalior"],
        ["2", "Bridging Machine Learning and Generative AI for Intelligent Vulnerability Detection", "Dr. Anant Nimkar et al.", "2026 IEEE IATMSI, Gwalior"],
        ["3", "Cryptographic API Misuse Detection Using BERT-Based Model", "Dr. Anant Nimkar et al.", "2026 IEEE ICITIIT, Kottayam"],
        ["4", "Comparative Analysis of Transformer Models for Skin Lesion Detection", "Dr. Disha Sail et al.", "2025 IEEE CVMI"],
        ["5", "NutriScan: Multimodal Deep Learning for Nutritional Deficiency Detection", "Dr. Disha Sail et al.", "3rd Int. Conf. on Computing & ML (CML 2026)"],
        ["6", "CyberVigilance: A Deep Learning Framework for Cyber Attack Forecasting", "Prof. Swapnali Kurhade et al.", "2025 IEEE INDISCON, NIT Rourkela"],
        ["7", "Handwriting Based Personality Assessment Using Image Processing", "Prof. Swapnali Kurhade et al.", "2025 IEEE ICDSINC, NIT Raipur"],
        ["8", "Interview Preparation Platform Using LLM, RAG and Speech Analysis", "Prof. Swapnali Kurhade et al.", "2026 ICAUC, Bengaluru & Shinawatra Univ."],
        ["9", "Exercise Posture Correction Using MediaPipe and Machine Learning", "Prof. Swapnali Kurhade, Prof. Reeta Koshy et al.", "2026 ICAUC, Bengaluru"],
        ["10", "Flood Risk Prediction Using Supervised & Unsupervised Machine Learning", "Ms. Falak Malik, Dr. Aarti Karande", "2025 IEEE INDISCON, NIT Rourkela"],
        ["11", "Multimodal Depression Detection System", "Dr. Nataasha Raul et al.", "3rd Int. Conf. on CS, Cyber Security & IT (CCIT-2025)"],
        ["12", "Real-Time IDS With Integrated Host and Network Monitoring Using ML", "Dr. Nataasha Raul et al.", "2025 IEEE ICDSINC, Raipur"],
        ["13", "LexChain: Blockchain-Based Solution to Safeguard Electronic Evidence", "Dr. Prasenjit Bhavathankar et al.", "2025 13th IEEE ISDFS, Boston, MA, USA"],
        ["14", "Generative Adversarial Networks for Image Augmentation: Architecture & Training", "Dr. Prasenjit Bhavathankar et al.", "2025 2nd IEEE APCIT, Mysore"],
        ["15", "Performance Analysis of ML Models for Advanced Persistent Threat Detection", "Prof. Azman Qureshi, Dr. D. Ambawade", "2025 5th IEEE CONIT, Hubballi"],
        ["16", "Beyond Hate Speech: Synthetic Conversational Dataset for Subtle Cyberbullying", "Prof. Sunil Ghane et al.", "2025 IEEE PuneCon, Pune"],
        ["17", "Finalyze: A RAG-Based Framework for Intelligent Financial Document Analysis", "Prof. Sunil Ghane et al.", "2025 5th IEEE ICERECT, Mandya"],
        ["18", "Dissecting Pain: Multi-Modal Study Using Advanced CNN & Transformer Models", "Prof. Sunil Ghane et al.", "2025 17th IEEE IC3, Noida"],
        ["19", "Comparative Study of Chunking Methods for Document Summarization", "Prof. Sunil Ghane et al.", "2025 8th IEEE ICCPCT, Kollam"],
        ["20", "Graph Neural Approaches to Credit Scoring: Interpretable Learning with GCN", "Dr. Jignesh Sisodia, Dr. K. Devadkar", "2025 2nd IEEE APCIT, Mysore"],
        ["21", "Uncertainty-Aware Osteoporosis Classification with Calibrated Transformer", "Dr. Jignesh Sisodia, Dr. K. Devadkar", "2025 7th IEEE ICIDCA, Coimbatore"],
        ["22", "KubeTranscode: Distributed FFmpeg System for DASH Transcoding at Scale", "Dr. Pramod Bide et al.", "2025 IEEE GIEST, Jamshedpur"],
        ["23", "Edu-Fy: Attendance Management & Attention Tracking Using Deep Learning", "Dr. Pramod Bide et al.", "2025 IEEE ICONAT, Goa"],
        ["24", "StreamScape: Adaptive Encoding & AI-Enhanced Subtitles and Thumbnails", "Dr. Pramod Bide et al.", "2025 IEEE GCON, Itanagar"],
        ["25", "Tutor-Aware Automated Essay Scoring with Knowledge Graph Interpretability", "Dr. Kailas Devadkar et al.", "2026 IEEE ICMI"]
    ]
    data_table(doc, ["Sr.", "Paper Title", "Author(s)", "Conference Venue & Details"], conf_papers, col_widths=[0.5, 2.7, 1.8, 2.0], compact=True)

    doc.add_page_break()

def build_section_12(doc):
    section_header(doc, "Faculty and Student Achievements", "SECTION 12 — HONOURS, AWARDS & HACKATHONS")

    data = load_docx_data()
    file12 = data['12. Faculty and Students Achievements.docx']

    sub_section_header(doc, "Faculty Achievements & Honours (AY 2025–26)", "12.1")
    fac_ach_table = file12['tables'][0]['rows']
    # Skip header
    fac_ach_rows = fac_ach_table[1:]
    data_table(doc, ["Faculty Member", "Distinction / Award / Recognition"], fac_ach_rows, col_widths=[2.4, 4.6])

    sub_section_header(doc, "Student Achievements & Competition Laurels", "12.2")
    body(doc, "Complete roster of all 75 national and international competitive achievements across FE, SE, TE, and BE classes:")

    stud_ach_table = file12['tables'][1]['rows']
    # Filter valid rows (skipping banner/category subheaders)
    stud_rows = []
    for r in stud_ach_table:
        if len(r) >= 5 and r[0].strip().isdigit():
            sr, event, name, org, cls = [c.strip() for c in r[:5]]
            stud_rows.append([sr, name, cls, event, org])

    data_table(doc, ["Sr.", "Student Name", "Class", "Event / Competition Details", "Awarding Organisation / Agency"], stud_rows, col_widths=[0.5, 1.8, 0.9, 2.2, 1.8], compact=True)

    doc.add_page_break()

def build_section_13(doc):
    section_header(doc, "Undergraduate Projects, M.Tech Dissertations & Ph.D.", "SECTION 13 — RESEARCH INVENTIVENESS")

    data = load_docx_data()
    file13 = data['13. List of Undergraduate Projects, M.Tech and PhD.docx']

    sub_section_header(doc, "B.Tech. Major Projects (Complete Departmental Roster)", "13.1")
    body(doc, "Complete roster of all 158 student researchers across 38 project groups working under faculty supervision:")

    ug_table = file13['tables'][0]['rows']
    ug_rows = []
    for r in ug_table[1:]:
        if len(r) >= 8 and r[0].strip():
            gr, sr, uid, sname, title, guide, domain, ptype = [c.strip() for c in r[:8]]
            ug_rows.append([gr, uid, sname, title, guide, domain, ptype])

    data_table(doc, ["Gr.", "UID", "Student Name", "Project Title", "Guide", "Domain", "Type"], ug_rows, col_widths=[0.5, 1.0, 1.4, 2.2, 1.2, 0.9, 0.7], compact=True)

    sub_section_header(doc, "M.Tech. Computer Engineering Dissertations", "13.2")
    body(doc, "Master of Technology dissertation research projects undertaken during AY 2025–26:")

    pg_table = file13['tables'][1]['rows']
    pg_rows = []
    for r in pg_table[1:]:
        if len(r) >= 5 and r[0].strip().isdigit():
            sn, uid, sname, mentor, title = [c.strip() for c in r[:5]]
            pg_rows.append([sn, uid, sname, mentor, title])

    data_table(doc, ["SN", "UID", "Student Name", "Faculty Mentor", "Dissertation Title"], pg_rows, col_widths=[0.5, 1.1, 1.8, 1.5, 2.3], compact=True)

    sub_section_header(doc, "Ph.D. Doctoral Research Program", "13.3")
    body(doc, "The Department of Computer Engineering is a recognized Research Centre of the University of Mumbai, with a sanctioned doctoral intake of 30 research scholars. Research domains include Distributed Systems, Cloud Security, Deep Learning Forensics, Computer Vision, and Natural Language Processing.")

    doc.add_page_break()
