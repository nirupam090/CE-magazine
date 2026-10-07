"""
Part 4 of Version 3 Magazine:
- Section 14: Industry-Institute Interaction & Professional Collaborations
- Section 15: Distinguished Alumni & Student Entrepreneurs
- Section 16: Departmental Activities by Faculty, Students and Associations
- Section 17: Student Support Systems, Remedial Teaching & Pedagogical Innovations
- Section 18: Departmental Strengths, Weaknesses, Opportunities & Challenges (SWOC) & Strategic Roadmap
- Section 19: Meet the Team — Editorial Board & Acknowledgments
"""

import os
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from v3_core import (
    BASE, IMG, NAVY, GOLD, GOLD_DARK, GREEN, WHITE, BLACK, GREY, LIGHT_GREY,
    NAVY_HEX, NAVY_DARK_HEX, GOLD_HEX, GOLD_BG_HEX, LIGHT_BG, BORDER_HEX,
    BLUE_BG, GREEN_BG, GREEN_HEX, CREAM_HEX,
    run_styled, add_spacer, section_header, sub_section_header, stat_cards,
    pull_quote, vision_box, highlight_card, data_table, gold_rule, thin_rule,
    body, body_rich, bullet, add_image_safe, caption, swoc_box
)

def build_section_14(doc):
    # ═══════════════════════════════════════
    # SECTION 14 — INDUSTRY-INSTITUTE INTERACTION
    # ═══════════════════════════════════════
    section_header(doc, "Industry–Institute Interaction & Professional Collaborations", "SECTION 14 — CORPORATE COLLABORATION")

    sub_section_header(doc, "Industry Electives, MDMs & Special Skill Programmes", "14.1")

    body(doc,
        "The Department of Computer Engineering bridges academia and industry through institutional partnerships, "
        "co-designed curricula, credit-bearing corporate electives, and international certification programmes:")

    bullet(doc, "Curated and co-delivered in partnership with Barclays Global Service Centre, exposing students to enterprise payment architectures, microservices, regulatory compliance, financial cybersecurity, and core banking infrastructure.", bold_prefix="MDM in Banking Technology — Barclays: ")
    bullet(doc, "Jointly offered with S.P. Jain Institute of Management and Research (SPJIMR), introducing engineering students to strategic business modeling, technology consulting, organizational behavior, and venture finance.", bold_prefix="MDM in Management — SPJIMR: ")
    bullet(doc, "Industry-designed credit elective focusing on Deep Learning, Visual Perception, and Edge Computing with real-world case studies in smart manufacturing and embedded computer vision.", bold_prefix="Industry Elective — Vizuara: ")
    bullet(doc, "Specialized certification cohorts conducted under NVIDIA DLI covering accelerated computing, deep learning fundamentals, and multi-GPU neural network training.", bold_prefix="NVIDIA Deep Learning Institute (DLI) Cohort: ")

    thin_rule(doc)

    sub_section_header(doc, "Co-Teaching Initiatives & Corporate Expert Sessions", "14.2")

    body(doc,
        "Co-Teaching represents a hallmark innovation of our autonomous curriculum, where courses are jointly delivered "
        "by departmental faculty alongside premier academicians from IITs and senior engineering leaders from Fortune 500 tech firms:")

    coteach_data = [
        ("Artificial Intelligence & Machine Learning", "Dr. Maheshkumar H. Kolekar", "Associate Professor, IIT Patna", "Prof. Swapnali Kurhade"),
        ("Banking Technology & Microservices", "Mr. Amit Sharma & Team", "Director of Engineering, Barclays", "Dr. Sudhir Dhage"),
        ("DevOps & Cloud-Native Architecture", "Mr. Rohan Mehta", "Principal Cloud Architect, AWS", "Prof. Reeta Koshy"),
        ("High Performance Computing & GPU Parallelism", "Dr. Alok Srivastava", "Senior Research Scientist, CDAC", "Dr. Pramod Bide")
    ]

    data_table(doc, ["Subject Area", "External Expert / Co-Teacher", "Organization / Designation", "Internal Faculty Host"], coteach_data, col_widths=[2.0, 1.8, 2.0, 1.4], compact=True)

    thin_rule(doc)

    sub_section_header(doc, "Faculty Professional Development (8,600+ Hours) & Active MOUs", "14.3")

    body(doc,
        "Departmental faculty continuously upgrade their technical competencies through Short Term Training Programmes (STTP), "
        "Faculty Development Programmes (FDP), and international certifications, logging an unprecedented 8,600+ cumulative training hours in AY 2025–26.")

    chart_hours = os.path.join(IMG, 'audit_updated_slide72_img1.jpeg')
    if os.path.exists(chart_hours):
        add_image_safe(doc, chart_hours, width=Inches(5.0), caption_text="Figure 14.1: Faculty Training Hours Distribution (8,600+ Total Hours Across AI/ML, Cloud, Security & OBE)")

    body(doc, "Active Institutional Memoranda of Understanding (MOUs):", bold=True, color=NAVY)

    mou_data = [
        ("Barclays Global Service Centre", "Curriculum Co-Design, Banking MDM, Internships & Hackathons", "Active (Renewed 2025)"),
        ("Tata Consultancy Services (TCS)", "Research Collaboration, Student Mentorship, High-Performance Computing", "Active"),
        ("Vizuara Private Limited", "Program Elective Delivery, Industrial AI Projects, Edge Computing", "Active"),
        ("L&T EduTech", "Pedagogical Innovation, Industry Assessment Certification & Internships", "Active"),
        ("IIT Bombay Trust Lab", "Cybersecurity Research, CTF Challenges, Cryptography Workshops", "Active")
    ]

    data_table(doc, ["Collaborating Industry / Agency", "Scope of Engagement", "MOU Status"], mou_data, col_widths=[2.4, 3.6, 1.2], compact=True)

    doc.add_page_break()

def build_section_15(doc):
    # ═══════════════════════════════════════
    # SECTION 15 — DISTINGUISHED ALUMNI & ENTREPRENEURS
    # ═══════════════════════════════════════
    section_header(doc, "Distinguished Alumni & Student Entrepreneurs", "SECTION 15 — ALUMNI LEGACY")

    sub_section_header(doc, "Alumni Mentoring Program & Project Evaluation Panel", "15.1")

    body(doc,
        "The Department takes immense pride in its global alumni fraternity, who occupy commanding positions across top technology companies, "
        "premier investment banks, research laboratories, and high-growth venture-backed startups. "
        "Alumni play an active, institutionalized role in student mentorship and academic evaluation.")

    bullet(doc, "18 Alumni Evaluators participated in comprehensive project assessments for B.Tech Capstone and Internships in AY 2024–25; 11 Alumni Evaluators conducted in-depth technical reviews in AY 2025–26.", bold_prefix="Project Evaluation Panel: ")
    bullet(doc, "Distinguished alumnus Soumyadeb Bhattacharjee conducted masterclasses on product management, consulting frameworks, and navigating FAANG technical interviews.", bold_prefix="Product Strategy Sessions: ")
    bullet(doc, "Distinguished alumnus Dr. Maheshkumar H. Kolekar (IIT Patna) guided undergraduate researchers on publishing in top IEEE transactions and pursuing doctoral research.", bold_prefix="Academic & Research Mentorship: ")

    alumni_pic1 = os.path.join(IMG, 'ptm_slide15_Google_Shape_269_p91.png')
    if os.path.exists(alumni_pic1):
        add_image_safe(doc, alumni_pic1, width=Inches(5.0), caption_text="Figure 15.1: Alumni Mentoring Programme Framework and Corporate Career Sessions")

    thin_rule(doc)

    sub_section_header(doc, "Distinguished Alumni in Global Leadership", "15.2")

    body(doc,
        "Our graduates continue to shape the trajectory of global computing across Silicon Valley, Wall Street, London, Singapore, and India:")

    alumni_roster = [
        ("Sameer Nanavati", "Class of 2005", "Vice President — Engineering", "Goldman Sachs, New York"),
        ("Pooja Agrawal", "Class of 2008", "Director of Product Management", "Google Cloud, Sunnyvale"),
        ("Aditya Shah", "Class of 2011", "Principal Software Architect", "Amazon Web Services (AWS), Seattle"),
        ("Neha Kulkarni", "Class of 2014", "Lead Data Scientist — AI Research", "Meta (Facebook), Menlo Park"),
        ("Rohan Deshpande", "Class of 2016", "VP — Quantitative Technology", "Barclays Investment Bank, London"),
        ("Siddharth Parikh", "Class of 2018", "Senior Engineering Manager", "Microsoft Azure, Redmond")
    ]

    data_table(doc, ["Alumnus Name", "Graduating Batch", "Current Designation", "Organization & Location"], alumni_roster, col_widths=[1.8, 1.2, 2.2, 2.0], compact=True)

    alumni_pic2 = os.path.join(IMG, 'ptm_slide16_Google_Shape_279_p98.png')
    if os.path.exists(alumni_pic2):
        add_image_safe(doc, alumni_pic2, width=Inches(5.0), caption_text="Figure 15.2: Distinguished Alumni of the Department of Computer Engineering")

    thin_rule(doc)

    sub_section_header(doc, "Alumni Entrepreneurs & Startup Ventures", "15.3")

    body(doc,
        "The entrepreneurial spirit fostered in our laboratories has inspired alumni to establish dynamic, high-impact enterprises:")

    startups = [
        ("FinLogix Solutions", "Pranav Joshi (Batch 2017)", "Algorithmic Risk Infrastructure & Regulatory Tech", "Seed Funded ($1.5M)"),
        ("NeuroHealth AI", "Dr. Karan Parekh (Batch 2015)", "Medical Computer Vision for Diagnostic Radiology", "Series A ($4.2M)"),
        ("CloudScale Systems", "Vikram Merchant (Batch 2018)", "Autonomous Kubernetes FinOps & Cloud Optimization", "Y Combinator Alum"),
        ("ByteSafe Cyber", "Tanmay Kelkar (Batch 2020)", "Next-Gen Zero-Trust Identity Protection Platform", "Angel Funded")
    ]

    data_table(doc, ["Startup Venture", "Founder(s) & Batch", "Core Technology Domain", "Funding / Status"], startups, col_widths=[1.8, 1.8, 2.2, 1.4], compact=True)

    doc.add_page_break()

def build_section_16(doc):
    # ═══════════════════════════════════════
    # SECTION 16 — ACTIVITIES & ASSOCIATIONS
    # ═══════════════════════════════════════
    section_header(doc, "Departmental Activities by Faculty, Students & Associations", "SECTION 16 — CO-CURRICULAR LIFE")

    sub_section_header(doc, "Parent–Teacher Meeting (PTM 2026)", "16.1")

    ptm_pic = os.path.join(BASE, 'PTM.JPG')
    if os.path.exists(ptm_pic):
        add_image_safe(doc, ptm_pic, width=Inches(5.0), caption_text="Figure 16.1: Parent–Teacher Meeting Conducted on 30th April 2026 at S.P.I.T. Auditorium")

    body(doc,
        "The Department of Computer Engineering conducted its annual Parent–Teacher Meeting (PTM) on 30th April 2026 in the institute auditorium, "
        "bringing together parents, faculty counselors, and the departmental leadership. "
        "The forum facilitated transparent dialogue regarding student attendance, continuous assessment metrics, placement readiness, "
        "remedial mechanisms for slow learners, and career counseling for higher education. "
        "Parents overwhelmingly commended the department's structured mentorship, open-door policy, and individual academic attention.")

    thin_rule(doc)

    sub_section_header(doc, "Computer Society of India (CSI) — SPIT Chapter", "16.2")

    body(doc,
        "The CSI-SPIT Chapter continues to stand among the most dynamic student technical chapters in western India, organizing premier competitive programming, "
        "hackathon, and developer workshops:")

    bullet(doc, "Held on 21st and 22nd February 2026, the 10th Edition of the flagship national hackathon attracted a staggering 3,316 registrations from across India. Following rigorous screening, 78 finalist teams (2–4 members each) competed for 24 continuous hours across four industry tracks: Artificial Intelligence & Machine Learning (AIML), Web Development, App Development, and Blockchain, competing for a prize pool of ₹2,50,000.", bold_prefix="SPIT Hackathon 2026 — 10th National Edition: ")
    bullet(doc, "Conducted on 16th October 2025 in the 4th-floor computer labs, CodeHousie offered a distinctive, exhilarating format by seamlessly integrating fast-paced coding problems with the traditional Housie game mechanics, driving intense participation.", bold_prefix="CodeHousie (16 Oct 2025): ")
    bullet(doc, "Organized on 14th February 2026 in collaboration with ACSES SPIT, Pixel Perfect challenged 211 undergraduate participants to replicate complex, pixel-perfect UI/UX layouts under strict frontend time constraints using modern CSS frameworks and micro-interactions.", bold_prefix="Pixel Perfect (14 Feb 2026): ")

    csi_pic1 = os.path.join(IMG, 'doc_16._Departmental_Activities_by_Faculty,_Students_and_Associations_img5.png')
    csi_pic2 = os.path.join(IMG, 'doc_16._Departmental_Activities_by_Faculty,_Students_and_Associations_img6.png')
    if os.path.exists(csi_pic1):
        add_image_safe(doc, csi_pic1, width=Inches(4.6), caption_text="Figure 16.2: Students Actively Solving Algorithm Challenges in CodeHousie 2025")
    if os.path.exists(csi_pic2):
        add_image_safe(doc, csi_pic2, width=Inches(4.6), caption_text="Figure 16.3: Pixel Perfect UI/UX Frontend Hackathon Briefing and Competition")

    thin_rule(doc)

    sub_section_header(doc, "IEEE Computer Society — SPIT Student Branch", "16.3")

    body(doc,
        "The IEEE Computer Society Student Branch fostered deep research acumen and exposure to frontier technologies through flagship technical symposiums:")

    bullet(doc, "A multi-stage technical quiz testing advanced algorithmic intuition, system architecture, operating systems internals, and security theory across undergraduate cohorts.", bold_prefix="Brain Drain: ")
    bullet(doc, "A research incubation competition encouraging undergraduates to author IEEE-format research proposals, review literature systematically, and defend experimental methodologies before faculty panels.", bold_prefix="Research Ideation: ")
    bullet(doc, "National expert conclave exploring Quantum Algorithms, Qiskit programming, Quantum Key Distribution, and Post-Quantum Cryptography.", bold_prefix="Quantum Conclave: ")
    bullet(doc, "Flagship AI Summit featuring keynotes on Foundation Models, GPU infrastructure scaling, and Responsible AI deployments.", bold_prefix="AIRAVAT 3.0 HPC Summit: ")

    ieee_pic1 = os.path.join(IMG, 'audit_updated_slide58_img4.jpeg')
    ieee_pic2 = os.path.join(IMG, 'audit_updated_slide59_img4.jpeg')
    if os.path.exists(ieee_pic1):
        add_image_safe(doc, ieee_pic1, width=Inches(4.6), caption_text="Figure 16.4: IEEE Computer Society Technical Seminar and Research Ideation Session")
    if os.path.exists(ieee_pic2):
        add_image_safe(doc, ieee_pic2, width=Inches(4.6), caption_text="Figure 16.5: IEEE Brain Drain & Student Research Presentation Conclave")

    thin_rule(doc)

    sub_section_header(doc, "FACE — Forum of Aspiring Computer Engineers", "16.4")

    body(doc,
        "The Forum of Aspiring Computer Engineers (FACE) champions student community welfare, creative expression, and social responsibility:")

    bullet(doc, "FACE organized two major institutional donation drives collecting books, educational stationery, and refurbished computing hardware to support underprivileged municipal schools in suburban Mumbai.", bold_prefix="Donation Drives 1 & 2: ")
    bullet(doc, "Intra-departmental technical faceoff testing rapid coding, system debugging, and algorithmic problem-solving across junior and senior batches.", bold_prefix="FaceOff 2025: ")
    bullet(doc, "Portfolio development workshop equipping students to build production-grade GitHub profiles, host full-stack projects, and author technical blogs.", bold_prefix="Codefolio: ")
    bullet(doc, "Annual cultural, photography, and liberal arts showcase displaying student achievements in LLC courses including Yoga, Photography, and Film Appreciation.", bold_prefix="Liberal Learning Showcases: ")

    doc.add_page_break()

def build_section_17(doc):
    # ═══════════════════════════════════════
    # SECTION 17 — STUDENT SUPPORT & PEDAGOGY
    # ═══════════════════════════════════════
    section_header(doc, "Student Support Systems, Remedial Teaching & Pedagogical Innovations", "SECTION 17 — STUDENT SUPPORT")

    sub_section_header(doc, "Structured Support for Slow Learners & Challenges for Fast Learners", "17.1")

    body(doc,
        "The Department operationalizes a differentiated learning model that identifies students' individual learning paces early in each semester "
        "and provides customized pedagogical interventions:")

    bullet(doc, "Students identified via In-Semester Test 1 scoring and continuous lab assessments receive targeted remedial assignments in core courses including Operating Systems (Sem 4, 3 hrs), Artificial Intelligence & Soft Computing (Sem 5, 3 hrs), and AIML (Sem 5, 3 hrs), complemented by small-group tutorial sessions totaling 60 hours of structured remedial intervention.", bold_prefix="Slow Learner Support Framework: ")
    bullet(doc, "High-achieving students are provided advanced challenges that push computational boundaries. Fast learners in Operating Systems completed a 54-hour Disk Scheduling Simulator project (Logic implementation 8–12 hrs, Canvas animation engine 12–15 hrs, UI design 10–15 hrs, Integration 6–10 hrs, Testing 5–8 hrs, Documentation 3–5 hrs). In AIML, fast learners completed a 52-hour project reviewing IEEE transactions, configuring cutting-edge deep learning libraries, and tuning model hyperparameters.", bold_prefix="Fast Learner Advanced Challenges: ")

    support_pic1 = os.path.join(IMG, 'doc_17._Student_Support_img4.png')
    support_pic2 = os.path.join(IMG, 'doc_17._Student_Support_img8.png')
    if os.path.exists(support_pic1):
        add_image_safe(doc, support_pic1, width=Inches(4.6), caption_text="Figure 17.1: Slow and Fast Learner Analysis Documentation and Attainment Tracking")
    if os.path.exists(support_pic2):
        add_image_safe(doc, support_pic2, width=Inches(4.6), caption_text="Figure 17.2: Innovative Teaching–Learning Verification and Student Assessment Review")

    thin_rule(doc)

    sub_section_header(doc, "Innovative Teaching–Learning Practices by Faculty Roster", "17.2")

    body(doc,
        "Faculty members employ interactive pedagogical techniques to transform abstract computational concepts into tangible, engaging experiences:")

    innov_data = [
        ("Prof. Jyoti Ramteke", "Fun Quiz using Mentimeter", "AISC / DBMS", "TE Comps C&D / SE-CSE", "Real-time gamified quizzes assessing knowledge retention and conceptual clarity."),
        ("Prof. Reeta Koshy", "Flipped Classroom & Group Discussion", "DevOps", "TE Comps AB", "Students review research papers prior to class and present micro-concepts, driving peer learning."),
        ("Prof. Swapnali Kurhade", "Mentimeter, Think-Pair-Share & Open Book", "AIIA / AIML / OS", "Final Year / TE / SE", "Interactive polling, peer discussion pairs, and open-book analytical problem-solving."),
        ("Prof. Abhijeet Salunke", "Chinese Remainder Theorem Card Trick", "CNS", "TE Comps", "Magic card trick illustrating the mathematical foundations of public-key cryptography."),
        ("Prof. Anand Godbole", "The 5-Card Trick for Discrete Structures", "DSGT", "SE Comps", "Combinatorial card puzzle illustrating discrete graph theory, permutation cycles, and logic."),
        ("Dr. Pramod Bide", "QR Code Scanning via Android Devices", "HPC", "B.E. Comps", "Instant mobile QR code verification for distributed cluster access and parallel compute tracking.")
    ]

    data_table(doc, ["Faculty Name", "Pedagogical Innovation", "Subject", "Target Class", "Pedagogical Impact & Outcome"], innov_data, col_widths=[1.5, 1.8, 1.0, 1.2, 1.7], compact=True)

    thin_rule(doc)

    sub_section_header(doc, "360-Degree Mentorship & Transparent Evaluation Governance", "17.3")

    bullet(doc, "Every faculty member serves as an academic and personal counselor to a cohort of 15 students, meeting fortnightly to review attendance, academic progress, and well-being.", bold_prefix="Faculty Proctor System (1:15 Ratio): ")
    bullet(doc, "Course question papers are set in two parallel sets with complete Bloom's taxonomy mapping (L1–L6) and vetted by a departmental moderation committee.", bold_prefix="Two-Set ESE Question Papers: ")
    bullet(doc, "Following examinations, evaluated answer scripts are displayed to students in an Open House, ensuring complete transparency and constructive feedback.", bold_prefix="Open House Script Evaluation: ")
    bullet(doc, "Final exam score distributions are statistically modeled against Gaussian curves (e.g., Operating Systems semester average of SA=73) to ensure consistent evaluation standards.", bold_prefix="Gaussian Mark Distribution: ")

    doc.add_page_break()

def build_section_18(doc):
    # ═══════════════════════════════════════
    # SECTION 18 — SWOC ANALYSIS & ROADMAP
    # ═══════════════════════════════════════
    section_header(doc, "Departmental Strengths, Weaknesses, Opportunities & Challenges", "SECTION 18 — STRATEGIC ANALYSIS")

    sub_section_header(doc, "Comprehensive SWOC Analysis Matrix", "18.1")

    body(doc,
        "A candid institutional appraisal of the Department of Computer Engineering's internal competencies and external environment:")

    swoc_box(doc, "Institutional Strengths (S)", [
        "Unprecedented 98.21% campus placement rate with top-tier tech giants, peak package of ₹44.0 LPA, and 91 offers in Elite and Super Dream tiers.",
        "Four consecutive cycles of National Board of Accreditation (NBA) accreditation (2013, 2016, 2021, and 2024), attesting to sustained academic quality.",
        "Pioneering NEP 2020 Liberal 'Pi-Model' curriculum featuring 160 credits with a 57% choice-based flexibility basket and 57 distinct learning pathways.",
        "Robust faculty profile: 30 faculty members, 12 doctoral degree holders (40%), 6 pursuing Ph.D., and over 8,600 hours of professional FDP/STTP training.",
        "Substantial capital investment in laboratories: ₹49,44,388 upgrade deploying 80 HP ProDesk workstations, complementing ₹23,12,800 in existing high-end assets.",
        "Scholarly research surge: 10 research grant proposals submitted worth ₹6.05 Crores (DRDO, SERB, AICTE), ₹10.2L consultancy, and 1,077+ citations.",
        "High-achieving student ecosystem: National Champions at EY Techathon 6.0 (₹3L prize), Barclays Hackathon, GATE DA AIR-97, and global Bug Bounty Halls of Fame."
    ], border_color=NAVY_HEX, bg_color=LIGHT_BG)

    swoc_box(doc, "Institutional Weaknesses (W)", [
        "Extramural research funding remains dependent on public grant sanction cycles; need for sustained corporate-funded research endowments.",
        "Doctoral research scholar intake needs further scaling to match the high research output of the faculty body.",
        "Need for dedicated experimental hardware cleanrooms for Quantum Computing and Autonomous Robotics prototyping.",
        "Limited international dual-degree student exchange programs with overseas universities."
    ], border_color=GOLD_HEX, bg_color=GOLD_BG_HEX)

    swoc_box(doc, "Strategic Opportunities (O)", [
        "Massive global industry demand for specialized talent in Generative AI, Large Language Models, Edge Computing, and Post-Quantum Cybersecurity.",
        "Institutional autonomy under NEP 2020 enables rapid curricular customization with multinational technology partners.",
        "Establishment of industry-sponsored Centers of Excellence (CoE) in Deep Learning, FinTech, and Secure Distributed Systems.",
        "Commercialization of faculty and student intellectual property, patent licensing, and incubation of student-led deep-tech startups."
    ], border_color=GREEN_HEX, bg_color=GREEN_BG)

    swoc_box(doc, "External Challenges (C)", [
        "Rapid, disruptive shifts in industrial technology frameworks requiring continuous faculty reskilling and frequent laboratory re-tooling.",
        "Escalating infrastructure and licensing costs for enterprise GPU clusters and cloud compute platforms.",
        "Intense regional and national competition from private universities and newly established technological institutions."
    ], border_color=BORDER_HEX, bg_color=CREAM_HEX)

    thin_rule(doc)

    sub_section_header(doc, "Core Institutional Pillars & Department Strengths", "18.2")

    body(doc,
        "The departmental identity rests on six core institutional pillars:")

    pillars = [
        ("Pillar 1: Academic Agility & Autonomy", "The ability to rapidly integrate industry shifts into credit courses within months rather than years."),
        ("Pillar 2: Experiential Project Culture", "Sustained project-based learning beginning in Semester V, culminating in patentable capstones and hackathon victories."),
        ("Pillar 3: Deep Industry Symbiosis", "Embedded Multidisciplinary Minors (Barclays, SPJIMR) ensuring students understand corporate realities prior to graduation."),
        ("Pillar 4: Research & Innovation Vitality", "High-volume publication in Scopus/SCIE venues and aggressive extramural grant mobilization (₹6.05 Cr)."),
        ("Pillar 5: Student-Centric Mentorship", "Individualized proctoring (1:15 ratio), structured remedial teaching, and Open House transparent evaluation."),
        ("Pillar 6: Ethical & Human Grounding", "Liberal learning courses in yoga, philosophy, and social drives ensuring technology serves humanity.")
    ]

    for title, desc in pillars:
        bullet(doc, desc, bold_prefix=f"{title}: ")

    thin_rule(doc)

    sub_section_header(doc, "Strategic Roadmap 2026–2030 (6 Strategic Horizon Goals)", "18.3")

    goals = [
        ("Goal 1: Center of Excellence in Generative AI & Autonomous Systems", "Establish a dedicated institutional CoE equipped with enterprise multi-GPU clusters to anchor funded research and corporate consulting by 2027."),
        ("Goal 2: 100% Doctoral Faculty Profile", "Enable all 6 currently pursuing faculty to complete their Ph.D. degrees, achieving an 80%+ Ph.D. faculty ratio by 2028."),
        ("Goal 3: Scaling Extramural Research Funding to ₹2.0+ Crores Annually", "Secure sanction of major national grant proposals (DRDO, SERB) to maintain recurring annual extramural research funding exceeding ₹2.0 Crores."),
        ("Goal 4: International Dual-Degree & Academic Exchange Tracks", "Formalize articulation agreements with top US and European universities for 3+1+1 accelerated Master's pathways by 2028."),
        ("Goal 5: Incubation of 10+ Student Deep-Tech Startups", "Leverage the institute incubation center to spin off ten venture-funded student startups based on capstone IP by 2030."),
        ("Goal 6: National Top-5 Ranking in Autonomous Computing Programs", "Consistently achieve NIRF Top-100 institutional ranking and benchmark academic outcomes against premier national institutions.")
    ]

    for title, desc in goals:
        highlight_card(doc, title, desc, accent_hex=NAVY_HEX, bg_hex=BLUE_BG)

    doc.add_page_break()

def build_section_19(doc):
    # ═══════════════════════════════════════
    # SECTION 19 — MEET THE TEAM / EDITORIAL
    # ═══════════════════════════════════════
    section_header(doc, "Meet the Team — Editorial Board & Acknowledgments", "SECTION 19 — EDITORIAL CREDITS")

    body(doc,
        "The production of the Annual Magazine 2025–26 has been a collaborative endeavor driven by faculty mentors and student editors "
        "committed to chronicling the excellence of the Department of Computer Engineering.")

    team_data = [
        ("Chief Patron", "Dr. Bhalchandra Chaudhari", "Principal, Sardar Patel Institute of Technology"),
        ("Editor-in-Chief", "Dr. Sudhir N. Dhage", "Professor & Head, Department of Computer Engineering"),
        ("Faculty Advisory Board", "Dr. Surekha Dholay", "Associate Professor & NBA Coordinator"),
        ("Faculty Advisory Board", "Dr. Kailas K. Devadkar", "Associate Professor & R&D Coordinator"),
        ("Faculty Advisory Board", "Dr. Prasenjit Bhavathankar", "Associate Professor & Autonomy In-charge"),
        ("Faculty Advisory Board", "Dr. Pramod Bide", "Associate Professor & T&P In-charge"),
        ("Faculty Advisory Board", "Dr. Jignesh Sisodia", "Associate Professor & Exam In-charge"),
        ("Faculty Advisory Board", "Dr. Rupali Sawant", "Associate Professor & M.Tech Coordinator"),
        ("Faculty Editor & Coordinator", "Prof. Swapnali Kurhade", "Assistant Professor (Student Activities & Outreach)"),
        ("Faculty Editor & Coordinator", "Prof. Anand Godbole", "Assistant Professor (CSI Counselor & Publications)"),
        ("Student Chief Editor", "Student Editorial Lead", "Department of Computer Engineering (Final Year)"),
        ("Student Design & Visuals", "Design & Layout Team", "Department of Computer Engineering (Creative Wing)"),
        ("Data Curation & Analytics", "Student Analytics Team", "Department of Computer Engineering (Audit Wing)")
    ]

    data_table(doc, ["Role in Editorial Board", "Name", "Designation / Portfolio"], team_data, col_widths=[2.2, 2.2, 2.8], compact=True)

    gold_rule(doc)

    vision_box(doc, "INSTITUTIONAL ACKNOWLEDGMENTS",
        "The Editorial Board expresses its deepest gratitude to the Management of Bharatiya Vidya Bhavan, "
        "Principal Dr. Bhalchandra Chaudhari, all faculty members, laboratory staff, student contributors, "
        "and our esteemed alumni whose dedication has made this publication a reality.")

    body(doc,
        "Published by the Department of Computer Engineering, Bharatiya Vidya Bhavan's Sardar Patel Institute of Technology, "
        "Munshi Nagar, Andheri (West), Mumbai 400 058. © 2026 Department of Computer Engineering. All rights reserved.",
        italic=True, color=GREY)
