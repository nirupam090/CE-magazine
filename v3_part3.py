"""
Part 3 of Version 3 Magazine:
- Section 08: Faculty and Staff Information
- Section 09 & 10: Faculty Additional Responsibilities & Departmental Governance
- Section 11: Consultancy, Grants, Patents and Research Publications
- Section 12: Faculty and Students Achievements
- Section 13: List of Undergraduate Projects, M.Tech Dissertations & Ph.D. Research
"""

import os
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from v3_core import (
    BASE, IMG, NAVY, GOLD, GOLD_DARK, GREEN, WHITE, BLACK, GREY, LIGHT_GREY,
    NAVY_HEX, NAVY_DARK_HEX, GOLD_HEX, GOLD_BG_HEX, LIGHT_BG, BORDER_HEX,
    BLUE_BG, GREEN_BG, GREEN_HEX,
    run_styled, add_spacer, section_header, sub_section_header, stat_cards,
    pull_quote, vision_box, highlight_card, data_table, gold_rule, thin_rule,
    body, body_rich, bullet, add_image_safe, caption, achievement_card
)

def build_section_08(doc):
    # ═══════════════════════════════════════
    # SECTION 08 — FACULTY AND STAFF INFORMATION
    # ═══════════════════════════════════════
    section_header(doc, "Faculty and Staff Information", "SECTION 08 — HUMAN CAPITAL")

    fac_pic = os.path.join(BASE, 'Faclty group pic.JPG')
    if os.path.exists(fac_pic):
        add_image_safe(doc, fac_pic, width=Inches(5.4), caption_text="Figure 8.1: Faculty and Staff Members of the Department of Computer Engineering, S.P.I.T. (AY 2025–26)")

    sub_section_header(doc, "Faculty Profile & Doctoral Strength", "8.1")

    stat_cards(doc, [
        ("30", "FACULTY MEMBERS", "navy"),
        ("12", "Ph.D. HOLDERS (40%)", "gold"),
        ("6", "PURSUING Ph.D. (20%)", "green"),
        ("1:15", "FACULTY-STUDENT RATIO", "navy")
    ])

    body(doc,
        "The Department of Computer Engineering boasts a stellar academic faculty comprising 30 dedicated educators, researchers, and mentors. "
        "With 12 holding doctoral degrees from premier national institutes (including IITs and VNIT) and 6 actively pursuing their Ph.D., "
        "the department blends profound academic scholarship with extensive practical industry acumen.")

    fac_roster = [
        ("1", "Dr. Sudhir N. Dhage", "Professor & HOD", "Ph.D., M.E., B.E.", "Distributed Systems, Security, Blockchain", "25 Years"),
        ("2", "Dr. Bhalchandra Chaudhari", "Principal & Professor", "Ph.D., M.Tech, B.E.", "Communication Networks, Signal Processing", "34 Years"),
        ("3", "Dr. Surekha Dholay", "Associate Professor", "Ph.D., M.E., B.E.", "Computer Vision, Machine Learning, Deep Learning", "22 Years"),
        ("4", "Dr. Kailas K. Devadkar", "Associate Professor", "Ph.D., M.Tech, B.E.", "Cloud Computing, Data Analytics, Natural Language Proc.", "21 Years"),
        ("5", "Dr. Prasenjit Bhavathankar", "Associate Professor", "Ph.D., M.Tech, B.E.", "Blockchain, Cybersecurity, Distributed Systems", "20 Years"),
        ("6", "Dr. Pramod Bide", "Associate Professor", "Ph.D., M.E., B.E.", "High Performance Computing, Big Data Analytics", "20 Years"),
        ("7", "Dr. Jignesh Sisodia", "Associate Professor", "Ph.D., M.E., B.E.", "Computer Architecture, Graph Neural Networks, FinTech", "19 Years"),
        ("8", "Dr. Rupali Sawant", "Associate Professor", "Ph.D., M.E., B.E.", "Explainable AI, Healthcare Informatics, Deep Learning", "18 Years"),
        ("9", "Dr. Nataasha Raul", "Assistant Professor", "Ph.D., M.E., B.E.", "Network Security, Intrusion Detection, Data Mining", "16 Years"),
        ("10", "Dr. Anant V. Nimkar", "Assistant Professor", "Ph.D., M.Tech, B.E.", "Cyber-Physical Systems, IoT, Augmented Reality", "15 Years"),
        ("11", "Dr. Disha Harish Sail", "Assistant Professor", "Ph.D., M.E., B.E.", "Image Processing, Deep Learning, Biomedical Computing", "14 Years"),
        ("12", "Dr. Renuka Pawar", "Assistant Professor", "Ph.D., M.E., B.E.", "Zero-Trust Security, Biometrics, Machine Learning", "13 Years"),
        ("13", "Prof. Anand Godbole", "Assistant Professor", "M.Tech, B.E. (Ph.D. Pursuing)", "Discrete Math, Machine Learning, Deep Models", "15 Years"),
        ("14", "Prof. Swapnali Kurhade", "Assistant Professor", "M.E., B.E. (Ph.D. Pursuing)", "AI in Industrial Engineering, Operating Systems, LLMs", "14 Years"),
        ("15", "Prof. Reeta Koshy", "Assistant Professor", "M.Tech, B.E. (Ph.D. Pursuing)", "Software Engineering, DevOps, Cloud Architecture", "14 Years"),
        ("16", "Prof. Abhijeet Salunke", "Assistant Professor", "M.E., B.E. (Ph.D. Pursuing)", "Cryptography, Network Security, Quantum Computing", "13 Years"),
        ("17", "Prof. Sunil B. Ghane", "Assistant Professor", "M.Tech, B.E. (Ph.D. Pursuing)", "Theory of Computation, NLP, Financial Analytics", "13 Years"),
        ("18", "Prof. Jyoti Ramteke", "Assistant Professor", "M.E., B.E. (Ph.D. Pursuing)", "Soft Computing, Database Systems, Computer Vision", "12 Years"),
        ("19", "Prof. Govind Gaundalkar", "Assistant Professor", "M.E., B.E.", "Digital Systems, Computer Organization, Embedded Systems", "12 Years"),
        ("20", "Prof. Sheetal Choudhary", "Assistant Professor", "M.Tech, B.E.", "Data Structures, Small Language Models, Web Technologies", "10 Years"),
        ("21", "Prof. Pallavi Ahya", "Assistant Professor", "M.E., B.E.", "Graph Neural Networks, Credit Scoring, Transformer Models", "9 Years"),
        ("22", "Prof. Falak Malik", "Assistant Professor", "M.E., B.E.", "Supervised Learning, Predictive Modeling, Risk Systems", "8 Years"),
        ("23", "Prof. Azman Qureshi", "Assistant Professor", "M.Tech, B.E.", "Advanced Persistent Threat Detection, Cyber Forensics", "8 Years"),
        ("24", "Prof. Varsha Hole", "Assistant Professor", "M.E., B.E.", "Object-Oriented Programming, Mobile Application Dev.", "7 Years"),
        ("25", "Prof. Sneha Deshmukh", "Assistant Professor", "M.Tech, B.E.", "Artificial Intelligence, Human-Computer Interaction", "6 Years"),
        ("26", "Prof. Nikhil Pardeshi", "Assistant Professor", "M.E., B.E.", "Algorithm Design, Cloud Computing, System Security", "6 Years"),
        ("27", "Prof. Priyanka Patil", "Assistant Professor", "M.Tech, B.E.", "Data Mining, Big Data Pipelines, Intelligent Systems", "5 Years"),
        ("28", "Prof. Amit Kadam", "Assistant Professor", "M.E., B.E.", "Network Protocols, Cybersecurity, Operating Systems", "5 Years"),
        ("29", "Prof. Rashmi Joshi", "Assistant Professor", "M.Tech, B.E.", "Machine Learning, Natural Language Processing", "4 Years"),
        ("30", "Prof. Tejaswini Patil", "Assistant Professor", "M.E., B.E.", "Web Architecture, Software Testing, Agile Engineering", "4 Years")
    ]

    data_table(doc, ["#", "Faculty Name", "Designation", "Highest Qualification", "Specialisation / Research Domain", "Exp."], fac_roster, col_widths=[0.4, 1.8, 1.3, 1.2, 1.8, 0.5], compact=True)

    sub_section_header(doc, "Technical, Laboratory & Administrative Supporting Staff", "8.2")

    staff_data = [
        ("1", "Mr. M. S. Parab", "Senior Laboratory Assistant", "Diploma in Computer Technology", "CCF & Network Infrastructure Maintenance"),
        ("2", "Mr. S. V. More", "Laboratory Assistant", "B.Sc. (Computer Science)", "Systems Architecture & AI Lab Support"),
        ("3", "Mrs. P. P. Kadam", "Laboratory Assistant", "Diploma in Information Technology", "Software Engineering & Cloud Computing Lab"),
        ("4", "Mr. A. R. Shinde", "Hardware / Network Technician", "Hardware & Networking Certified", "CCN Lab Cisco Rack Maintenance & Campus LAN"),
        ("5", "Mr. R. B. Jadhav", "Laboratory Attendant", "Higher Secondary", "Laboratory Inventory & Physical Asset Stewardship"),
        ("6", "Mrs. S. S. Sawant", "Department Administrative Executive", "B.Com, Office Automation", "Department Records, Academic Files, BoS Coordination")
    ]

    data_table(doc, ["#", "Staff Member Name", "Designation", "Qualification", "Key Responsibilities"], staff_data, col_widths=[0.4, 1.8, 1.6, 1.4, 1.8], compact=True)

    doc.add_page_break()

def build_section_09_and_10(doc):
    # ═══════════════════════════════════════
    # SECTION 09 & 10 — RESPONSIBILITIES & GOVERNANCE
    # ═══════════════════════════════════════
    section_header(doc, "Faculty Additional Responsibilities & Governance", "SECTIONS 09 & 10 — INSTITUTIONAL ROLES & BODIES")

    sub_section_header(doc, "Additional Responsibilities Entrusted to the Faculty", "9.0")

    body(doc,
        "Beyond classroom instruction and research supervision, faculty members shoulder pivotal institutional portfolios, "
        "ensuring that departmental operations adhere to highest standards of academic governance, accreditation compliance, and student welfare.")

    resp_data = [
        ("Dr. Sudhir N. Dhage", "Head of Department (HOD)", "Chairman BoS, DAPC, Departmental Administration & Resource Allocation"),
        ("Dr. Surekha Dholay", "NBA / NAAC Department Coordinator", "Outcome-Based Education Quality Assurance, Accreditation Compliance"),
        ("Dr. Kailas K. Devadkar", "Department R&D & Ph.D. Coordinator", "Research Proposals, Doctoral Committee In-charge, Patents & Grants"),
        ("Dr. Prasenjit Bhavathankar", "Institutional Autonomy In-charge", "Curriculum Structure Framework, NEP 2020 Implementation In-charge"),
        ("Dr. Pramod Bide", "Department Training & Placement In-charge", "Corporate Relations, Campus Recruiter Interface, Placement Audits"),
        ("Dr. Jignesh Sisodia", "Examination & Moderation In-charge", "Question Paper Moderation, Continuous Assessment & ESE Evaluation"),
        ("Dr. Rupali Sawant", "M.Tech Programme Coordinator", "Postgraduate Academics, Dissertation Progress Reviews, PG Quality"),
        ("Dr. Nataasha Raul", "Alumni Relations & Outreach In-charge", "Alumni Mentorship Network, Industry Evaluator Coordination"),
        ("Prof. Anand Godbole", "CSI SPIT Student Branch Counselor", "Computer Society of India Activities, Hackathons & Tech Events"),
        ("Prof. Swapnali Kurhade", "IEEE Computer Society Branch Counselor", "IEEE Technical Conclaves, International Student Activities"),
        ("Prof. Reeta Koshy", "Department Timetable & Academic In-charge", "Classroom/Lab Scheduling, Faculty Workload Optimization"),
        ("Prof. Abhijeet Salunke", "FACE Staff Advisor & Student Welfare", "Forum of Aspiring Computer Engineers, Cultural & Social Drives")
    ]

    data_table(doc, ["Faculty Member", "Administrative Portfolio", "Scope of Institutional Responsibility"], resp_data, col_widths=[1.8, 2.0, 3.2], compact=True)

    thin_rule(doc)

    sub_section_header(doc, "Departmental Governance: Advisory Board, Board of Studies & DAPC", "10.0")

    body(doc,
        "The Department operates under a robust statutory governance framework ensuring continuous academic relevance, curriculum rigor, and stakeholder alignment:")

    highlight_card(doc, "Departmental Advisory Board (DAB)",
        "Composed of eminent academicians from IIT Bombay, VNIT Nagpur, and VJTI, alongside senior technology leaders from "
        "Tata Consultancy Services, Barclays, and Morgan Stanley. The DAB convenes annually to review strategic direction, "
        "infrastructure investments, and long-term research trajectories.",
        accent_hex=NAVY_HEX, bg_hex=BLUE_BG)

    highlight_card(doc, "Board of Studies in Computer Engineering (BoS)",
        "The statutory body responsible for approving syllabus revisions, new course introductions, multidisciplinary minors, "
        "and evaluation schemes. Includes subject experts from Mumbai University, industry subject-matter experts, and distinguished alumni.",
        accent_hex=GOLD_HEX, bg_hex=GOLD_BG_HEX)

    highlight_card(doc, "Departmental Academic Performance Committee (DAPC)",
        "Monitors teaching-learning delivery, course attainment metrics, student attendance, slow/fast learner interventions, "
        "and continuous assessment transparency on a monthly basis.",
        accent_hex=GREEN_HEX, bg_hex=GREEN_BG)

    doc.add_page_break()

def build_section_11(doc):
    # ═══════════════════════════════════════
    # SECTION 11 — RESEARCH, GRANTS, PATENTS & PUBLICATIONS
    # ═══════════════════════════════════════
    section_header(doc, "Consultancy, Grants, Patents and Research Publications", "SECTION 11 — RESEARCH & SCHOLARSHIP")

    sub_section_header(doc, "Faculty Citation Metrics & Research Impact", "11.1")

    stat_cards(doc, [
        ("1,077+", "TOTAL CITATIONS", "navy"),
        ("14", "HIGHEST h-INDEX", "gold"),
        ("18", "HIGHEST i10-INDEX", "green"),
        ("40", "PUBLICATIONS (2025–26)", "navy")
    ])

    body(doc,
        "The faculty members of the Department of Computer Engineering maintain an active research profile indexed across Google Scholar, Scopus, and Web of Science. "
        "The scholarly work spans Distributed Computing, Cybersecurity, Generative AI, Explainable AI, Healthcare Diagnostics, and Blockchain Architecture.")

    cit_data = [
        ("1", "Dr. Sudhir N. Dhage", "285", "11", "13", "Distributed Security, Cloud Integrity, Blockchain"),
        ("2", "Dr. Bhalchandra Chaudhari", "340", "14", "18", "Communication Networks, Signal Processing"),
        ("3", "Dr. Surekha Dholay", "145", "8", "7", "Medical Image Analysis, Deep Learning, Vision"),
        ("4", "Dr. Kailas K. Devadkar", "112", "7", "6", "Natural Language Processing, Cloud Analytics"),
        ("5", "Dr. Prasenjit Bhavathankar", "95", "6", "5", "Blockchain Networks, Smart Contracts"),
        ("6", "Dr. Pramod Bide", "88", "6", "4", "High-Performance Computing, Parallel Systems"),
        ("7", "Dr. Jignesh Sisodia", "74", "5", "4", "Computer Architecture, Graph Neural Nets"),
        ("8", "Dr. Rupali Sawant", "68", "5", "4", "Explainable AI, Tuberculosis Healthcare Models"),
        ("9", "Dr. Nataasha Raul", "52", "4", "3", "Host & Network IDS, Machine Learning"),
        ("10", "Dr. Anant V. Nimkar", "48", "4", "3", "IoT Systems, Cyber-Physical Testbeds"),
        ("11", "Prof. Anand Godbole", "38", "3", "2", "Discrete Mathematics, Deep Learning"),
        ("12", "Prof. Swapnali Kurhade", "35", "3", "2", "Cyber Attack Forecasting, Industrial AI")
    ]

    data_table(doc, ["#", "Faculty Name", "Total Citations", "h-index", "i10-index", "Primary Research Domain"], cit_data, col_widths=[0.4, 2.0, 1.1, 0.8, 0.8, 2.1], compact=True)

    thin_rule(doc)

    sub_section_header(doc, "Research Grant Proposals (10 Proposals Worth ₹6.05 Crores) & Consultancy", "11.2")

    body(doc,
        "In Academic Year 2025–26, the department achieved a major milestone in extramural research mobilization by submitting 10 competitive research grant proposals "
        "totaling ₹6.05 Crores to premier national funding agencies including DRDO, SERB CRG, AICTE RPS, and RGSTC. "
        "Furthermore, Dr. Anant Nimkar secured and executed a high-impact industrial consultancy project valued at ₹10.2 Lakhs.")

    grants_table = [
        ("1", "AI-Driven Real-Time Vulnerability Analysis for Defense Networks", "DRDO (Defense Research & Dev. Org.)", "₹85,00,000", "Dr. Sudhir Dhage (PI)"),
        ("2", "Interpretable Graph Neural Networks for Precision Cardiovascular Diagnostics", "SERB - Core Research Grant (CRG)", "₹68,50,000", "Dr. Surekha Dholay (PI)"),
        ("3", "Decentralized Zero-Trust Identity Federation using Post-Quantum Cryptography", "SERB - CRG", "₹62,00,000", "Dr. Prasenjit Bhavathankar (PI)"),
        ("4", "Voice-Enabled Interactive Multimodal AI for Assistive Neurological Care", "AICTE - Research Promotion Scheme (RPS)", "₹25,00,000", "Dr. Kailas Devadkar (PI)"),
        ("5", "Distributed Scalable Video Transcoding on Heterogeneous Edge Clusters", "AICTE - RPS", "₹22,50,000", "Dr. Pramod Bide (PI)"),
        ("6", "Uncertainty-Aware Deep Learning Framework for Bone Density Osteoporosis Scoring", "RGSTC (Rajiv Gandhi Sci & Tech)", "₹45,00,000", "Dr. Jignesh Sisodia (PI)"),
        ("7", "MEDxplain: Multi-Center XAI Framework for Non-Invasive TB & Pulmonary Lesions", "SERB - CRG", "₹72,00,000", "Dr. Rupali Sawant (PI)"),
        ("8", "Host-Network Integrated Threat Hunting Engine using Real-Time Transformers", "AICTE - RPS", "₹24,00,000", "Dr. Nataasha Raul (PI)"),
        ("9", "Physics-Guided Generative AI for Thermal-Aware Embedded Edge IoT Devices", "DRDO", "₹95,00,000", "Dr. Anant Nimkar (PI)"),
        ("10", "Synthetic Speech Forensics and Deepfake Audio Attribution Framework", "RGSTC", "₹1,06,00,000", "Prof. Swapnali Kurhade (PI) / Dr. Dhage")
    ]

    data_table(doc, ["#", "Proposal Title", "Funding Agency", "Outlay (₹)", "Principal Investigator"], grants_table, col_widths=[0.4, 2.7, 1.8, 1.0, 1.3], compact=True)

    chart_funding = os.path.join(IMG, 'audit_updated_slide42_img4.jpeg')
    if os.path.exists(chart_funding):
        add_image_safe(doc, chart_funding, width=Inches(5.0), caption_text="Figure 11.1: Submitted Research Grant Proposals (₹6.05 Crores) Across Premier National Funding Bodies")

    highlight_card(doc, "Industry Consultancy Milestone — ₹10.2 Lakhs",
        "Dr. Anant V. Nimkar successfully led an enterprise consultancy project valued at ₹10,20,000 for developing an "
        "Industrial IoT Asset Telemetry and Predictive Maintenance Platform for manufacturing automation partners, "
        "exemplifying direct translation of departmental research into commercial industry workflows.",
        accent_hex=GREEN_HEX, bg_hex=GREEN_BG)

    thin_rule(doc)

    sub_section_header(doc, "Intellectual Property & Patents Filed", "11.3")

    patents = [
        ("Patent 1 (Published)", "SOUND2SKETCH: AN AI POWERED INTERACTIVE VOICE-ENABLED DRAWING TOOL",
         "Dr. Kailas Devadkar, Satyam Jaiswal, Siddhant Rao, Dr. Jignesh Sisodia, Komal Tarachandani",
         "Application No.: 20 2025 104 311 | Date of Filing: 25/07/2025 | Date of Publication: 20/08/2025",
         "An assistive AI system translating spoken descriptive prompts into real-time semantic sketches and vectorized illustrations for educational and accessibility applications."),
        ("Patent 2 (Filed)", "AI-DRIVEN REAL-TIME AUDITORY & SENSOR MONITORING SYSTEM FOR ENVIRONMENT SAFETY",
         "Prof. Anand Godbole, Dr. Surekha Dholay",
         "Indian Patent Application No. 202521089412 | Date of Filing: 14/11/2025",
         "An edge-computing embedded architecture that processes ambient sound patterns and vibration telemetry using TinyML models for early industrial anomaly detection.")
    ]

    for p_tag, title, inventors, filing_info, desc in patents:
        bullet(doc, f"{title}\nInventors: {inventors}\n{filing_info}\nScope: {desc}", bold_prefix=f"{p_tag}: ")

    thin_rule(doc)

    sub_section_header(doc, "International Journal & Conference Publications (Scopus / SCIE / IEEE)", "11.4")

    body(doc,
        "During Academic Year 2025–26, the department published 40 research papers: 15 in reputed international journals (including 9 Scopus and 4 SCIE indexed) "
        "and 25 in peer-reviewed IEEE and international conference proceedings.")

    chart_research = os.path.join(IMG, 'audit_updated_slide71_img1.jpeg')
    if os.path.exists(chart_research):
        add_image_safe(doc, chart_research, width=Inches(5.0), caption_text="Figure 11.2: Department Research Mix — Journals (Scopus/SCIE), International Conferences & Patents")

    pub_sample = [
        ("1", "Physics-Based AR Framework for Realistic Home Furnishing", "A. Dhara, N. Gode, G. Gangurde, Dr. A. V. Nimkar", "IEEE IATMSI 2026", "DOI: 10.1109/IATMSI68868.2026"),
        ("2", "Bridging ML and GenAI for Source Code Vulnerability Remediation", "A. Deokar et al., Dr. A. V. Nimkar", "IEEE IATMSI 2026", "DOI: 10.1109/IATMSI68868.2026"),
        ("3", "NutriScan: Hybrid Deep Learning for Medical Deficiency Image Analysis", "R. Sharma, Dr. R. Sawant, T. Sinkar, P. Tayshete", "CML 2026 (Springer)", "SMU Sikkim, March 2026"),
        ("4", "CyberVigilance: Deep Learning Framework for Cyber Attack Forecasting", "A. Agarwal, M. Ayalwar, R. Baviskar, Prof. S. Kurhade", "IEEE INDISCON 2025", "DOI: 10.1109/INDISCON66021.2025"),
        ("5", "Interview Prep Platform using LLM, RAG and Speech Analysis", "Y. Gajewar, M. Lodha, V. Parmar, Prof. S. Kurhade", "ICAUC 2026", "Bengaluru, Jan 2026"),
        ("6", "LexChain: Blockchain-Based Solution to Safeguard Legal E-Evidence", "N. Adhikari, M. Chhajed, D. Raichura, Dr. P. Bhavathankar", "IEEE ISDFS 2025", "Boston, USA, DOI: 10.1109/ISDFS"),
        ("7", "Beyond Hate Speech: Synthetic Dataset for Subtle Cyberbullying", "A. Jangam, U. Avaiya, H. Sawarkar, Prof. S. Ghane", "IEEE PuneCon 2025", "DOI: 10.1109/PuneCon67554.2025"),
        ("8", "Graph Neural Approaches to Credit Scoring with GCN and GAT", "P. Ahya, Dr. J. Sisodia, Dr. K. Devadkar", "IEEE APCIT 2025", "DOI: 10.1109/APCIT65661.2025"),
        ("9", "KubeTranscode: Distributed FFmpeg System for Scalable DASH Transcoding", "S. Reddy, O. Sarode, H. Singh, Dr. P. Bide", "IEEE GIEST 2025", "DOI: 10.1109/GIEST66547.2025"),
        ("10", "Real-Time IDS with Integrated Host & Network Monitoring", "S. Nandwani, D. Otari, A. Tripathi, Dr. N. Raul", "IEEE ICDSINC 2025", "DOI: 10.1109/ICDSINC66221.2025")
    ]

    data_table(doc, ["#", "Paper Title", "Authors", "Conference / Journal Venue", "Citation / Identifier"], pub_sample, col_widths=[0.4, 2.5, 1.8, 1.3, 1.2], compact=True)

    doc.add_page_break()

def build_section_12(doc):
    # ═══════════════════════════════════════
    # SECTION 12 — FACULTY AND STUDENTS ACHIEVEMENTS
    # ═══════════════════════════════════════
    section_header(doc, "Faculty and Students Achievements", "SECTION 12 — HONORS & RECOGNITION")

    sub_section_header(doc, "Faculty Honors, NPTEL Certifications & Pedagogical Awards", "12.1")

    fac_achievements = [
        ("Prof. Anand Godbole", "NPTEL Discipline Star Award + Elite Certifications in 'Joy of Computing using Python' & 'Theory of Computation'"),
        ("Dr. Pramod Bide", "Best Innovative Teaching Award conferred by L&T EduTech for exemplary integration of mobile cloud learning"),
        ("Dr. Jignesh Sisodia", "Best Software Project Award conferred by Cantilever + Elite NPTEL Certification in Computer Architecture"),
        ("Prof. Govind Gaundalkar", "Best Question Paper Award for the subject Digital Systems (Autonomous Examination Division)"),
        ("Dr. Jignesh Sisodia & Prof. Abhijeet Salunke", "Best Question Paper Award for Computer Communication and Networks (CCN)"),
        ("Prof. Reeta Koshy", "Elite NPTEL Certification in Software Engineering (Top 5% candidate nationally)"),
        ("Prof. Swapnali Kurhade", "Elite NPTEL Certification in 'AI in Industrial and Management Engineering'"),
        ("Prof. Sunil Ghane", "NPTEL National Certification in Theory of Computation"),
        ("Dr. Surekha Dholay", "Invited Technical Program Committee (TPC) Reviewer for IEEE Flagship International Conferences")
    ]

    data_table(doc, ["Faculty Member", "National Honor / Pedagogical Distinction"], fac_achievements, col_widths=[2.4, 4.8], compact=True)

    thin_rule(doc)

    sub_section_header(doc, "Student National Hackathons & Innovation Crowns", "12.2")

    body(doc,
        "Students of the Department of Computer Engineering demonstrated exceptional problem-solving and software engineering mastery on the national stage, "
        "clinching victories at premier corporate and collegiate hackathons:")

    # Embed Winner Photos (Grid of 6)
    w_imgs = [
        ('audit_updated_slide60_img4.jpeg', "EY Techathon 6.0 National Winners — ₹3,00,000 Cash Prize"),
        ('audit_updated_slide60_img6.jpeg', "Barclays National Hackathon Champion Team"),
        ('audit_updated_slide60_img7.jpeg', "Google Cloud GenAI Exchange Hackathon Winners"),
        ('audit_updated_slide60_img8.jpeg', "GATE DA 2026 All India Rank 97 — Atharva Jagtap"),
        ('audit_updated_slide60_img9.jpeg', "Adappt 4.0 Hackathon 1st Place Winners"),
        ('audit_updated_slide60_img10.jpeg', "Ethereum & Zoho Global Bug Bounty Hall of Fame — Madhav Shah")
    ]

    for fname, cap in w_imgs:
        p_path = os.path.join(IMG, fname)
        if os.path.exists(p_path):
            add_image_safe(doc, p_path, width=Inches(3.8), caption_text=cap)

    achievement_card(doc, "EY Techathon 6.0 — National Champion (1st Place)",
        "Soham Ghorpade, Dikshant Karande, Dwayne Fernandes, Swain Dias", "Ernst & Young (EY)", "₹3,00,000 CASH PRIZE", GOLD_HEX)

    achievement_card(doc, "GenAI Exchange Hackathon — 3rd Runner Up",
        "Daksh Chetan Jain", "Google Cloud", "NATIONAL RECOGNITION", NAVY_HEX)

    achievement_card(doc, "Barclays National Hackathon — First Prize",
        "Divyanshu Sharma, Varad Kulkarni", "Barclays Global Service Centre", "PPO + ₹1,00,000 AWARD", GREEN_HEX)

    achievement_card(doc, "ICPC Asia Amritapuri — Regionalist Finalists",
        "Raj Kale, Abhimanyu Kadhane, Adityavardhan Iyengar", "International Collegiate Programming Contest", "ASIA REGIONALIST", NAVY_HEX)

    achievement_card(doc, "IIT Bombay Techfest AI Eduthon — Winner",
        "Daksh Jain", "Cograd / IIT Bombay Techfest", "FIRST PRIZE", GOLD_HEX)

    achievement_card(doc, "APOGEE Advanced Tech Hackathon — Winner",
        "Soham Budhewar, Pratik Jambhule, Kartik Kharat", "BITS Pilani", "FIRST PRIZE", GREEN_HEX)

    achievement_card(doc, "Adappt 4.0 Hackathon — 1st Place Winner",
        "Madhav Shah", "NMIMS MPSTME", "FIRST POSITION", NAVY_HEX)

    achievement_card(doc, "Nomura KakushIN Hackathon — Finalist Honors",
        "Tanush Salian, Suryansh Singh", "Nomura Holdings", "EXCELLENCE AWARD", GOLD_HEX)

    thin_rule(doc)

    sub_section_header(doc, "GATE DA AIR-97, Bug Bounty Halls of Fame & Cohort Roster", "12.3")

    highlight_card(doc, "GATE DA 2026 — All India Rank 97 (Score: 816)",
        "Atharva Jagtap (BE Computer Engineering) secured an extraordinary All India Rank 97 in the newly introduced "
        "GATE Data Science and Artificial Intelligence paper with a standardized score of 816 out of 1000, "
        "ranking among the top 0.1% candidates nationally.",
        accent_hex=GOLD_HEX, bg_hex=GOLD_BG_HEX)

    highlight_card(doc, "Global Security Research & Bug Bounty Halls of Fame",
        "Madhav Shah (TE Computer Engineering) earned induction into the prestigious Ethereum Foundation Bug Bounty Hall of Fame "
        "and Zoho Security Hall of Fame for responsibly disclosing critical smart contract and web application security vulnerabilities. "
        "He also received Certificates of Appreciation from GeeksforGeeks for critical security patches.",
        accent_hex=GREEN_HEX, bg_hex=GREEN_BG)

    body(doc, "Representative Student Achievements Across Cohorts (AY 2025–26):", bold=True, color=NAVY)

    cohort_achievements = [
        ("FE", "Nivaan Shah", "Shri Ram Investment Competition (D-Street, SRCC)", "1st Position Nationally"),
        ("SE", "Rudraksh Shah", "National Case Study Competition (SRCC Delhi)", "1st Position"),
        ("SE", "Amit Gandhi, Stavan Desai, Piyush Fatnani, Yash Dharamshi", "Recursion 7.0 Hackathon (RGIT)", "2nd Runner Up"),
        ("SE", "Nitin Gaikwad", "Equinox 26 Technical Fest (KKWIEER)", "1st Position"),
        ("SE", "Mayank Khadayate", "Enigma Hackathon & Ignite26 Pitching Competition", "Winner & Runner Up"),
        ("TE", "Sanjana Pillai, Aditya Khabiya, Harsh Malpani", "Chaitanya National Case Competition (IIM Indore)", "1st Position"),
        ("TE", "Aniruddh Iyer", "Break The Case (BITS Hyderabad) & Vishleshan (IIM Raipur)", "Winner & Runner Up"),
        ("TE", "Soham Kher", "Marketing Showdown (SJMSOM, IIT Bombay)", "National Winner"),
        ("TE", "Sahil Gupta, Vinayak Yadav, Nikhil Vishwakarma", "India Innovates 2k26 (Municipal Corp. of Delhi)", "Top Innovation Award"),
        ("TE", "Vir Bhalani, Daksh Jain, Vishwesh Nair, Jill Mistry", "HackFusion International Hackathon (IEEE SPIT/UMIT)", "1st Place Winner"),
        ("BE", "Soorya Sivaramakrishnan", "NPTEL Optimization Theory & Algorithms", "Elite + Silver Medal"),
        ("BE", "Muskan Sharma", "HSBC CTF Hackathon 2025 (IIT Bombay Trust Lab)", "Special Mention Finalist")
    ]

    data_table(doc, ["Class", "Student Name(s)", "Event / Competition Details", "Achievement Level"], cohort_achievements, col_widths=[0.6, 2.2, 2.6, 1.8], compact=True)

    doc.add_page_break()

def build_section_13(doc):
    # ═══════════════════════════════════════
    # SECTION 13 — UNDERGRADUATE PROJECTS & RESEARCH
    # ═══════════════════════════════════════
    section_header(doc, "List of Undergraduate Projects, M.Tech & Ph.D. Research", "SECTION 13 — RESEARCH PROJECTS & INNOVATIONS")

    sub_section_header(doc, "Capstone Project Domain Analytics & Research Breadth", "13.3")

    body(doc,
        "The B.Tech Major Capstone Projects represent the culmination of undergraduate training, requiring teams of 2 to 4 students "
        "to conceptualize, architect, implement, and benchmark complex computational systems. In AY 2025–26, 57 project groups (158 students) "
        "executed capstones across diverse cutting-edge domains.")

    chart_domains = os.path.join(IMG, 'audit_updated_slide20_img4.jpeg')
    if os.path.exists(chart_domains):
        add_image_safe(doc, chart_domains, width=Inches(4.8), caption_text="Figure 13.1: Final Year Capstone Project Domain Distribution — AI/ML (48%), Web/Cloud (24%), Security/Blockchain (16%), IoT/Embedded (12%)")

    highlight_card(doc, "NTRO Defense Research Internship Spotlight",
        "National Technical Research Organisation (NTRO), Government of India: Final-year student researchers were selected for "
        "prestigious national security computing internships, developing specialized algorithms in signal intelligence, "
        "cryptanalysis, and automated anomaly detection for defense applications.",
        accent_hex=NAVY_HEX, bg_hex=BLUE_BG)

    thin_rule(doc)

    sub_section_header(doc, "Representative B.Tech Major Capstone Projects (AY 2025–26)", "13.1")

    btech_projects = [
        ("Gr. 1", "AI-Driven Automated Financial Portfolio Management System", "Atharva Bambere, Darshan Bardeo, Arjavi Kakulte", "Dr. B. N. Chaudhari", "Machine Learning", "Application"),
        ("Gr. 2", "AI-Driven Financial Health Evaluation & Wealth Advisory Framework", "Maureen Miranda, Sara Nainwar, Ayushi Japsare", "Dr. Sudhir Dhage", "Machine Learning", "Application"),
        ("Gr. 3", "Comparative Study of Deep & Hybrid Models for Alzheimer's MRI", "Shardul Mahindrakar, Krrish Nichanii, Piyush Mehta", "Dr. Sudhir Dhage", "Deep Learning", "Research"),
        ("Gr. 4", "AI-Powered Adaptive Learning & Skill Diagnostic Platform", "Kanishka Amritkar, Lokesh Chaudhari", "Dr. Sudhir Dhage", "Artificial Intelligence", "Application"),
        ("Gr. 5", "Dynamic Real-Time Route Optimization for Urban Navigation", "Sneha Sambare, Suhani Garate, Sanika Sule", "Dr. Surekha Dholay", "Machine Learning", "Research"),
        ("Gr. 6", "Deepfake Detection using Attention Fusion of XceptionNet & DCT", "Shrinivas Waney, Atharva Sankhe", "Dr. Surekha Dholay", "Deep Learning", "Research"),
        ("Gr. 7", "FundChain: Transparent Blockchain Fundraising using Solidity", "Manish Jadhav, Mayur Solankar, Vishesh Savani", "Dr. P. Bhavathankar", "Blockchain", "Application"),
        ("Gr. 8", "BharatNFT: Scalable Marketplace for Digital Cultural Assets", "Shesh Shrikant, Ronit Chintrate, Varsha Bojja", "Dr. P. Bhavathankar", "Blockchain", "Application"),
        ("Gr. 9", "Mindease: Multimodal Stress Detection and Relaxation System", "Harsh Shah, Tej Shah, Veer Shah", "Prof. Anand Godbole", "Artificial Intelligence", "Application"),
        ("Gr. 10", "Injury Detection and Biomechanical Prevention for Football Players", "Daksh Jain, Bhaskar Kapase, Shreyas Naik", "Prof. Anand Godbole", "Deep Learning", "Research"),
        ("Gr. 11", "AI-Powered Automated Legal Contract Analysis for Indian Law", "Venisha Kalola, Hrishikesh Dalal, Dhureen Shettigar", "Dr. Kailas Devadkar", "Artificial Intelligence", "Research"),
        ("Gr. 12", "SAMARTH: Secure Multimodal Anomaly Recognition using Blockchain", "Vishal Chaure, Team", "Dr. Kailas Devadkar", "Blockchain", "Research"),
        ("Gr. 13", "MEDxplain: Explainable AI in Tuberculosis Diagnosis", "Devika Nikam, Pushkar Mhatre, Vanashree Tajane", "Dr. Rupali Sawant", "Artificial Intelligence", "Research"),
        ("Gr. 14", "Hierarchical Acoustic Monitoring for Real-Time Forest Surveillance", "Omkar Surve, Darsh Tulsiyan, Akhil Vaidya", "Dr. Sudhir Dhage", "Machine Learning", "Research"),
        ("Gr. 15", "Adaptive Quantization of Medical Vision Transformers using RL", "Yash Desai, Arya Mehta, Krishna Sharma", "Dr. P. Bhavathankar", "Deep Learning", "Research")
    ]

    data_table(doc, ["Gr.", "Project Title", "Students Team", "Faculty Guide", "Domain", "Type"], btech_projects, col_widths=[0.5, 2.3, 1.8, 1.2, 0.8, 0.6], compact=True)

    thin_rule(doc)

    sub_section_header(doc, "M.Tech Dissertation Research Projects (AY 2025–26)", "13.2")

    mtech_projects = [
        ("1", "Ameya Nile Ganesh", "Dr. Sudhir Dhage", "ML-Based Food Waste Reduction and Smart Restaurant Management System"),
        ("2", "Barkur Tushar Shyam", "Dr. Jignesh Sisodia", "Explainable Cross-Domain User Preference Modeling using Transformer Embeddings"),
        ("3", "Dalvi Vaishnavi Ramesh", "Prof. Reeta Koshy", "ClaimSage: An Agentic RAG Framework for Deep Reasoning over Insurance Claims"),
        ("4", "Deshpande Vedang P.", "Dr. Aarti Karande", "Reinforcement Learning for Multi-Agent Energy Optimization in Microgrids"),
        ("5", "Dhande Suyog Suresh", "Dr. D. R. Kalbande", "Explainable AI for Legal Case Prioritization using BERT-Based Word Embeddings"),
        ("6", "Jadhao Om Sandip", "Prof. Jyoti Ramteke", "Image Question Answering for Ethical Education: Semantic & Moral Framework"),
        ("7", "Korgaonkar Gauri V.", "Dr. Renuka Pawar", "Emotion-Aware Zero-Trust Continuous Access Control using Multimodal Biometrics"),
        ("8", "Lagade Sudarshan A.", "Prof. Abhijeet Salunke", "Spatio-Temporal Deep Learning for National Water Demand Prediction in India"),
        ("9", "Mahant Durgesh S.", "Dr. Nataasha Raul", "Smart Navigation System Integrating Real-Time Hazard Detection & Routing"),
        ("10", "Maharana Kiran K.", "Dr. P. Bhavathankar", "Spatiotemporal Graph Neural Networks for Water Contamination Risk Forecasting"),
        ("11", "Malik Falak Yusuf", "Dr. Pramod Bide", "Emoji-Free Sarcasm Detection in Code-Mixed Text via Emotion-Polarity Divergence"),
        ("12", "Morye Amey Dilip", "Dr. Sujata Kulkarni", "Yoga Mitra: Multilingual Framework for Real-Time Pose Correction & Assistance"),
        ("13", "Parab Devang Kishor", "Dr. Kailas Devadkar", "AI-Powered Mental Health Prediction with Integrated Psychiatric Assistance"),
        ("14", "Shaikh Fazil Ahmed", "Prof. Sheetal Choudhary", "Decomposing Recall and Reasoning Failures in Quantized Small Language Models"),
        ("15", "Sole Kshitija Mahesh", "Prof. Sunil Ghane", "AgriXplain: Explainable Region-Aware Ensemble System for Crop Yield Forecasting")
    ]

    data_table(doc, ["#", "Postgraduate Scholar", "Faculty Mentor", "Master's Thesis Research Title"], mtech_projects, col_widths=[0.4, 1.8, 1.6, 3.4], compact=True)

    doc.add_page_break()
