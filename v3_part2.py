"""
Part 2 of Version 3 Magazine:
- Section 05: About the Department — Full Curriculum Architecture & Syllabus Highlights
- Section 06: Laboratory Facilities, Smart Classrooms & Major Equipment Added
- Section 07: Overview of Academic Programs, Placements, Internships & Higher Studies
"""

import os
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from v3_core import (
    BASE, IMG, NAVY, GOLD, GOLD_DARK, GREEN, WHITE, BLACK, GREY, LIGHT_GREY,
    NAVY_HEX, NAVY_DARK_HEX, GOLD_HEX, GOLD_BG_HEX, LIGHT_BG, BORDER_HEX,
    run_styled, add_spacer, section_header, sub_section_header, stat_cards,
    pull_quote, vision_box, highlight_card, data_table, gold_rule, thin_rule,
    body, body_rich, bullet, add_image_safe, caption, company_tags
)

def build_section_05(doc):
    # ═══════════════════════════════════════
    # SECTION 05 — ABOUT THE DEPARTMENT & CURRICULUM
    # ═══════════════════════════════════════
    section_header(doc, "About the Department — Curriculum and Syllabus Highlights", "SECTION 05 — ACADEMIC ARCHITECTURE")

    sub_section_header(doc, "Redefining Engineering Education: The Liberal 'Pi-Model' Architecture", "5.1")

    pull_quote(doc,
        "Engineering education is undergoing a fundamental transformation. The engineer of today is expected to do much more than master technical concepts. "
        "Along with strong domain knowledge, students need the ability to think critically, communicate effectively, work across disciplines, "
        "understand emerging technologies, solve real-world problems and continuously adapt to a rapidly changing world.",
        "Curriculum Framework Manifesto")

    body(doc,
        "Recognising this changing landscape, Sardar Patel Institute of Technology has introduced its Liberal 'Pi-Model' of Engineering Education under its autonomous curriculum framework. "
        "The curriculum is designed not merely to produce technically competent engineers, but to nurture well-rounded professionals with intellectual, emotional, physical and spiritual dimensions of development. "
        "The curriculum structure provides students with a strong foundation in mathematics, basic sciences and engineering sciences while progressively moving towards programme-specific knowledge, "
        "electives, experiential learning, multidisciplinary education, research and industry engagement.")

    body(doc,
        "One of the defining features of the new curriculum is flexibility and choice. The curriculum follows a Choice Based Credit System (CBCS) in which students are provided multiple opportunities "
        "to shape their academic journey according to their interests and aspirations. The curriculum features a 57% choice-based component, making the learning pathway remarkably flexible and student-centric. "
        "The 160-credit structure distributes learning across Basic and Engineering Sciences, Programme Core Courses, Programme Electives, Experiential Learning, Open Electives and Skill Enhancement Courses. "
        "Programme-related courses account for 55% of the curriculum, while substantial credit space is dedicated to electives and experiential learning. "
        "This flexibility allows students to move beyond a common, fixed academic pathway and gradually develop an academic profile aligned with their individual interests.")

    body(doc,
        "Building Strong Foundations: The initial semesters focus on developing a strong engineering foundation. Courses such as Engineering Calculus, Differential Equations and Complex Analysis, "
        "Linear Algebra, and Probability and Statistics, along with engineering science courses, establish the mathematical and scientific base required for advanced engineering study. "
        "For Computer Engineering students, this foundation is complemented by courses in Digital Systems and Microprocessors, Programming, Data Structures, Database Management Systems, "
        "Computer Organization and Architecture, and other core computing areas. The curriculum therefore follows a progressive structure: students first develop foundational knowledge, "
        "then acquire core computing competencies, and subsequently move towards advanced areas and specialised domains.")

    body(doc,
        "From Core Computing to Emerging Technologies: The Computer Engineering curriculum reflects the rapidly changing technological landscape. "
        "By the third and fourth semesters, students encounter core subjects such as Discrete Structures and Graph Theory, Database Management Systems, Computer Organization and Architecture, "
        "Data Structures, Computer Communications and Networks, Operating Systems, Software Engineering, and Design and Analysis of Algorithms. "
        "The fifth semester takes students further into contemporary computing through Theory of Computation, Artificial Intelligence and Machine Learning, Cryptography and Network Security, "
        "and Distributed Computing, together with an open elective and Mini Project I. This progression is particularly significant because it connects fundamental concepts directly with the technologies shaping modern computing.")

    body(doc,
        "A Strong Emphasis on Artificial Intelligence and Emerging Domains: Artificial Intelligence is not treated as an isolated subject in the curriculum. "
        "Students are introduced to Artificial Intelligence and Machine Learning as part of the programme core and can subsequently specialise through an extensive programme elective basket. "
        "The elective basket includes cutting-edge domains such as:")

    ai_electives = [
        "Deep Learning & Neural Architectures",
        "Generative AI & Large Language Models",
        "Explainable and Responsible AI (XAI)",
        "Agentic AI & Multi-Agent Frameworks",
        "Data Mining & Knowledge Discovery",
        "Natural Language Processing & Speech",
        "Big Data Analytics & Streaming Engines",
        "Decision Making & Business Intelligence",
        "Quantum Computing & Quantum Algorithms",
        "Quantum Machine Learning (QML)"
    ]
    company_tags(doc, ai_electives)

    body(doc,
        "This structure gives students an opportunity to move from foundational AI knowledge towards specialised and emerging areas. "
        "For students interested in cybersecurity, the curriculum provides another specialised pathway covering Cyber Security and Ethical Hacking, Blockchain, Digital Forensics, and Information and System Security. "
        "Similarly, students interested in hardware and intelligent systems can explore Embedded Systems, IoT and Edge Computing, Robotics and Automation, and Embedded AI and TinyML. "
        "The curriculum also provides pathways in Computer Vision, AR/VR, Visual Computing, High Performance Computing, Cloud Computing, Edge Computing, and Wireless Ad-hoc and Sensor Networks.")

    thin_rule(doc)

    sub_section_header(doc, "Specialisation Pathways, Multidisciplinary Minors & Experiential Learning", "5.2")

    body(doc,
        "Learning Beyond the Boundaries of Computer Engineering: Modern technological problems rarely belong to a single discipline. "
        "Recognising this reality, the curriculum provides a significant multidisciplinary component through Open Electives and Multidisciplinary Minor (MDM) tracks. "
        "Students can explore areas such as Industrial IoT, Signal Processing and Communication, Computer Engineering, Interface and Experience Design, Artificial Intelligence and Machine Learning, "
        "Financial and Strategic Management, Banking Technology, Entrepreneurship and Innovation, English Literature, Mathematics, Physics, and Systems Engineering. "
        "This creates an opportunity for a Computer Engineering student to understand technology in a broader context—for example, combining computing with finance, management, physics, entrepreneurship or user experience design. "
        "Such exposure is increasingly important in an era where successful technological solutions require an understanding of users, organisations, markets and society alongside technical implementation.")

    highlight_card(doc, "Experiential Learning Architecture: From Concepts to Creation",
        "A major feature of the curriculum is its emphasis on Experiential Learning. Students undertake Mini Project I in Semester V and "
        "Mini Project II in Semester VI, followed by a Main Capstone Project and semester-long research or industry internships.\n"
        "This progression creates an organic movement across the learning continuum:\n"
        "LEARN  →  EXPLORE  →  BUILD  →  EXPERIMENT  →  SOLVE  →  CREATE\n"
        "Instead of restricting project work to the final year, project-based learning is introduced much earlier, empowering students to convert theoretical knowledge into working systems.",
        accent_hex=GOLD_HEX, bg_hex=GOLD_BG_HEX)

    body(doc,
        "Research as a Pathway, Not Just an Outcome: For students with a strong inclination towards research, the curriculum provides a dedicated Honours by Research pathway. "
        "Students can undertake structured research stages during summer terms and subsequent semesters: Research Stage I, Research Stage II, Research III, and Research IV. "
        "The scheme specifies that students earning 18 additional credits through the prescribed research stages, along with meeting the stipulated CGPA requirement of 8.25, "
        "become eligible for the prestigious 'Honours by Research' degree pathway. This creates an academic environment where research becomes a sustained undergraduate experience.")

    body(doc,
        "Industry and Research Exposure: The curriculum provides dedicated pathways for Research / Industry Internships during the later semesters. "
        "Students utilize experiential learning credits to gain immersion in corporate and premier laboratory environments, applying theoretical concepts to live industry codebases.")

    body(doc,
        "Engineering with a Human Dimension: Technology does not exist in isolation from society. The curriculum deliberately includes Humanities, Social Sciences, Management, and Liberal Learning components. "
        "Students choose from courses covering microeconomics, law for engineers, psychology, Indian history, political science, philosophy, and foreign languages. "
        "The Liberal Learning Course (LLC) basket goes even further, offering: Yoga, Introduction to Bhagavad Geeta, Film Appreciation, First Aid, Photography, Innovation and Entrepreneurship, "
        "Interpersonal Effectiveness and Team Building, and Martial Arts. These courses nurture personal discipline, cultural awareness, and responsible citizenship.")

    body(doc,
        "Skills for the Professional World: Substantial emphasis is placed on Skill Enhancement Courses including Design Thinking, Engineering Exploration, Professional Communication, "
        "Corporate Communication, Data Science Laboratory, Real-Life Problem Solving in AI, and Computing System and Application Development. "
        "These courses hone the critical abilities to communicate ideas crisply, approach ambiguity systematically, and engineer polished solutions.")

    body(doc,
        "Multiple Pathways, One Engineering Identity & Multiple Exit Credentials: Every student possesses distinct career aspirations. "
        "One student may dive deeply into AI, another into cybersecurity, a third into financial technology, and another into tech entrepreneurship. "
        "Simultaneously, the curriculum incorporates NEP 2020 multi-entry and multi-exit provisions: students completing year one earn a Certificate in Engineering; "
        "year two, a Diploma in Engineering; year three, a B.Sc. in Engineering; and year four, the full B.Tech. in Computer Engineering degree.")

    stat_cards(doc, [
        ("160 CR", "DEGREE CREDITS", "navy"),
        ("57%", "CHOICE COMPONENT", "gold"),
        ("18 CR", "HONOURS RESEARCH", "green"),
        ("4 CR", "LIBERAL LEARNING", "navy")
    ])

    doc.add_page_break()

def build_section_06(doc):
    # ═══════════════════════════════════════
    # SECTION 06 — LAB FACILITIES & EQUIPMENT
    # ═══════════════════════════════════════
    section_header(doc, "Laboratory Facilities, Smart Classrooms & Major Equipment Added", "SECTION 06 — INFRASTRUCTURE & FACILITIES")

    sub_section_header(doc, "State-of-the-Art Department Laboratories & Smart Classrooms", "6.1")

    body(doc,
        "The Department of Computer Engineering maintains world-class laboratory infrastructure designed to provide undergraduate and postgraduate students "
        "with an optimal environment for computational experimentation, applied research, and advanced software engineering. "
        "Spread across the 4th, 5th, and 6th floors of the institution, our laboratories feature specialized hardware, gigabit networking, uninterruptible power systems, "
        "and modern ergonomic layouts.")

    labs_info = [
        ("Central Computing Facility (CCF)", "Houses high-density desktop clusters used for foundation programming, algorithm design, operating systems, and large-scale competitive programming contests."),
        ("Computer Communication & Networks Lab (CCN - 6th Floor)", "Equipped with enterprise Cisco Catalyst switches, routers, protocol analyzers, network simulation software, and hardware testbeds for cybersecurity and distributed computing."),
        ("Project & Research Innovation Laboratory", "Dedicated space for final-year B.Tech capstone teams and M.Tech dissertation researchers, equipped with dedicated high-end workstations and IoT development hardware."),
        ("Data Science & Artificial Intelligence Laboratory", "GPU-accelerated compute machines with NVIDIA RTX architecture supporting deep learning model training, computer vision pipelines, and NLP research experiments."),
        ("Systems Architecture & Microprocessor Lab", "Features hardware emulation kits, digital logic analyzers, microcontrollers, FPGA boards, and IoT edge kits for embedded systems and hardware-software co-design."),
        ("Software Engineering & Cloud Computing Lab", "Configured for Agile sprints, full-stack web and mobile application engineering, automated testing, containerisation (Docker, Kubernetes), and CI/CD pipelines.")
    ]

    for title, desc in labs_info:
        bullet(doc, desc, bold_prefix=f"{title}: ")

    add_spacer(doc, 4)

    # Embed Lab Images
    lab_img1 = os.path.join(IMG, 'doc_6._Lab_facilities,_Smart_Classroom_img4.png')
    lab_img2 = os.path.join(IMG, 'doc_6._Lab_facilities,_Smart_Classroom_img10.png')
    if os.path.exists(lab_img1):
        add_image_safe(doc, lab_img1, width=Inches(4.8), caption_text="Figure 6.1: Modernized High-Performance Computing Laboratory at S.P.I.T. Department of Computer Engineering")
    if os.path.exists(lab_img2):
        add_image_safe(doc, lab_img2, width=Inches(4.8), caption_text="Figure 6.2: Interactive Smart Classroom with 4K Touch Display and Lecture Capture Integration")

    body(doc,
        "Smart Classrooms: All departmental lecture halls are equipped as Smart Classrooms featuring interactive 75-inch 4K touchscreen panels, "
        "integrated lecture capture facilities, ceiling-mounted array microphones, high-fidelity acoustic reinforcement, and seamless Wi-Fi 6 wireless coverage. "
        "These classrooms facilitate hybrid instruction, flipped-classroom discussions, interactive Mentimeter polls, and real-time live demonstrations.")

    thin_rule(doc)

    sub_section_header(doc, "Major Equipment & Hardware Modernisation in 2025–26", "6.2")

    body(doc,
        "In Academic Year 2025–26, the Department underwent a massive institutional hardware upgrade backed by ₹49,44,388 in capital expenditure, "
        "significantly augmenting the existing equipment base previously invested (₹23,12,800).")

    stat_cards(doc, [
        ("₹49,44,388", "2025–26 NEW UPGRADE", "navy"),
        ("80 UNITS", "HP PRODESK 400 G7 PCS", "gold"),
        ("₹23,12,800", "PRIOR / CONCURRENT ASSETS", "green"),
        ("₹72.57L", "TOTAL LAB VALUATION", "navy")
    ])

    equip_data = [
        ("HP ProDesk 400 G7 Microtower Workstations", "80 Units", "Intel Core i5-10500 (10th Gen), 16GB DDR4 RAM, 512GB NVMe M.2 SSD, 19.5\" IPS LED Monitor, Windows 11 Pro", "Computer Labs 1 & 2", "₹49,44,388"),
        ("Enterprise Cisco Catalyst Networking Suite", "4 Switches / Routers", "Cisco Catalyst 2960-X Gigabit managed switches, 24-port PoE+, structured CAT6A patch panels, rack mounts", "CCN Lab (6th Fl.)", "₹5,80,000"),
        ("Dell PowerEdge Department Server Rack", "2 Tower Servers", "Intel Xeon Silver 4210, 64GB ECC RAM, 4TB SAS RAID 5, redundant hot-plug power supplies for departmental virtualization", "Server Room", "₹6,50,000"),
        ("NVIDIA RTX Deep Learning Workstations", "4 Workstations", "Intel Core i9, 32GB RAM, NVIDIA GeForce RTX 3080/4070 (12GB VRAM), dual 27\" 4K IPS monitors for AI training", "Project & AI Lab", "₹7,20,000"),
        ("Embedded IoT & Edge AI Prototyping Kits", "25 Kits", "Raspberry Pi 4 (8GB), NVIDIA Jetson Nano Dev Kits, ESP32 modules, sensor arrays, camera modules", "Embedded Systems Lab", "₹2,10,000"),
        ("Interactive Smart Classroom Display Units", "2 Panels", "ViewSonic 75\" 4K Ultra HD Interactive Flat Panels with multi-touch stylus, OPS slot-in PC modules, mobile trolleys", "Smart Classrooms", "₹1,52,800")
    ]

    data_table(doc, ["Equipment Description", "Qty", "Technical Specifications", "Deployment Location", "Total Cost"], equip_data, col_widths=[1.8, 0.7, 2.3, 1.2, 1.0], compact=True)

    doc.add_page_break()

def build_section_07(doc):
    # ═══════════════════════════════════════
    # SECTION 07 — ACADEMIC PROGRAMS & PLACEMENTS
    # ═══════════════════════════════════════
    section_header(doc, "Overview of Academic Programs, Placements, Internships & Higher Studies", "SECTION 07 — STUDENT ACADEMICS & CAREERS")

    sub_section_header(doc, "Academic Cohorts (UG, PG, Ph.D.) & Statistical Enrollment", "7.1")

    body(doc,
        "The Department of Computer Engineering offers premier degree programs recognized for rigorous academic standards, high admission cutoff percentiles, "
        "and outstanding graduate outcomes. S.P.I.T.'s autonomous curriculum attracts top engineering aspirants from across Maharashtra and India.")

    program_data = [
        ("B.Tech in Computer Engineering", "Undergraduate", "4 Years", "180 + 36 (DSE) = 216", "864 Students", "MHT-CET / JEE Main (>99.4%ile)"),
        ("M.Tech in Computer Engineering", "Postgraduate", "2 Years", "18", "36 Students", "GATE (Computer Science)"),
        ("Ph.D. in Computer Engineering", "Doctoral", "3–5 Years", "As per vacancy", "22 Research Scholars", "PET / GATE + Research Interview")
    ]

    data_table(doc, ["Degree Programme", "Level", "Duration", "Annual Intake", "Total Enrollment", "Admission Mechanism"], program_data, col_widths=[1.8, 1.0, 0.8, 1.3, 1.1, 1.5], compact=True)

    body(doc,
        "Academic Outcomes Snapshot: The department achieved an extraordinary 99.1% graduation rate for the final-year cohort. "
        "Analysis of academic results indicates consistent distinctions and first-class degrees across semesters, supported by strong internal mentoring and remedial support.")

    chart_outcomes = os.path.join(IMG, 'audit_updated_slide73_img1.jpeg')
    if os.path.exists(chart_outcomes):
        add_image_safe(doc, chart_outcomes, width=Inches(5.0), caption_text="Figure 7.1: Department Academic Outcomes Snapshot — 71% Placed, 23% Higher Studies, 6% Entrepreneurship/Competitive Paths")

    thin_rule(doc)

    sub_section_header(doc, "Campus Placements 2025–26 (98.21% Placement Rate)", "7.2")

    body(doc,
        "The 2025–26 placement season for the Department of Computer Engineering established new benchmarks despite broader tech industry fluctuations. "
        "A total of 110 students were placed out of 112 eligible candidates, registering an extraordinary placement percentage of 98.21%. "
        "The recruitment drive witnessed participation from global financial institutions, top-tier tech giants, consulting majors, and dynamic AI startups.")

    stat_cards(doc, [
        ("98.21%", "PLACEMENT SUCCESS RATE", "green"),
        ("₹44.0 LPA", "HIGHEST DOMESTIC CTC", "gold"),
        ("₹14.85 LPA", "AVERAGE PACKAGE", "navy"),
        ("₹13.50 LPA", "MEDIAN SALARY", "navy")
    ])

    body(doc, "Salary Tier Distribution for 2025–26 Placements:", bold=True, color=NAVY)

    tiers = [
        ("Elite Tier (≥ 20.0 LPA)", "21 Offers", "PhonePe, Morgan Stanley, BlackRock, JP Morgan Chase, Zeta, BrowserStack"),
        ("Super Dream Tier (12.0 – 20.0 LPA)", "70 Offers", "Barclays (21), Nomura (6), IKS Health (5), ICICI Lombard (5), General Mills (4), Baker Hughes (4), We Matter (4)"),
        ("Dream Tier (8.0 – 12.0 LPA)", "10 Offers", "Loginext (3), HSBC (3), Wissen Technology (2), IDfy (2)"),
        ("Normal Tier (< 8.0 LPA)", "9 Offers", "Core technical and enterprise IT services companies")
    ]

    data_table(doc, ["Placement Salary Tier", "Offers Count", "Representative Recruiting Companies"], tiers, col_widths=[2.4, 1.2, 3.4], compact=True)

    chart_salary = os.path.join(IMG, 'audit_updated_slide61_img8.jpeg')
    chart_placements = os.path.join(IMG, 'audit_updated_slide24_img1.jpeg')
    if os.path.exists(chart_salary):
        add_image_safe(doc, chart_salary, width=Inches(5.0), caption_text="Figure 7.2: Salary Tier Distribution for Computer Engineering (21 Elite, 70 Super Dream, 10 Dream, 9 Normal)")
    if os.path.exists(chart_placements):
        add_image_safe(doc, chart_placements, width=Inches(5.0), caption_text="Figure 7.3: Department Overall Placement Statistics and Historic Recruitment Trends")

    body(doc, "Complete Roster of Recruiting Companies (Academic Year 2025–26):", bold=True, color=NAVY)

    recruiters_table = [
        ("Barclays", "21", "Nomura", "6"),
        ("BlackRock", "5", "IKS Health", "5"),
        ("ICICI Lombard", "5", "General Mills", "4"),
        ("Baker Hughes", "4", "We Matter", "4"),
        ("PhonePe", "3", "JP Morgan & Chase", "3"),
        ("Loginext", "3", "HSBC", "3"),
        ("Wissen Technology", "2", "IDfy", "2"),
        ("Russell Investments", "2", "Zeta", "2"),
        ("UBS", "2", "Nouryon", "2"),
        ("Bombay Stock Exchange (BSE)", "2", "Revvity", "2"),
        ("ISS Stoxx", "2", "EightGen AI", "2"),
        ("Morgan Stanley", "1", "IDFC First Bank", "1"),
        ("WorkIndia", "1", "Auxo AI", "1"),
        ("BrowserStack", "1", "Hyperverge", "1"),
        ("Mahindra Finance", "1", "Equity Data Science", "1"),
        ("Dolat Capital", "1", "Quantiphi", "1"),
        ("TransUnion", "1", "KPMG", "1"),
        ("CashRich", "1", "OM Partners", "1"),
        ("CarTrade", "1", "Oracle Financial (OFSS)", "1"),
        ("Infosys", "1", "Embeddos", "1"),
        ("Versor Investments", "1", "GEP Worldwide", "1"),
        ("Sure Financial", "1", "HERE Technologies", "1"),
        ("Quantanite", "1", "Unlox Technologies", "1"),
        ("OmneNEST Technology", "1", "NXPEC Technologies", "1"),
        ("Disha Clip Fitting", "1", "Finseal Software", "1"),
        ("Communioun", "1", "Proximity IT Solutions", "1")
    ]

    data_table(doc, ["Company Name", "Count", "Company Name", "Count"], recruiters_table, col_widths=[2.4, 0.8, 2.4, 0.8], compact=True)

    thin_rule(doc)

    sub_section_header(doc, "Industry Internships (Batch 2022–2026) & Higher Studies Progression", "7.3")

    body(doc,
        "Every student undergoes a mandatory semester-long industry or research internship during their final year. "
        "For the graduating batch of 2022–2026, 77 students secured highly competitive industry internships with corporate stipends reaching up to ₹1,00,000 per month. "
        "These internships served as launchpads for Pre-Placement Offers (PPOs), with over 65% of interned students converting their internships into permanent positions.")

    intern_summary = [
        ("Barclays Global Service Centre", "18 Students", "₹60,000 / month", "Core Banking, Cloud Platform, DevOps"),
        ("Nomura Holdings", "6 Students", "₹75,000 / month", "Capital Markets, Quantitative Trading, Analytics"),
        ("BlackRock Investment Mgmt", "5 Students", "₹85,000 / month", "Aladdin Platform, Fixed Income Analytics"),
        ("Morgan Stanley", "2 Students", "₹1,00,000 / month", "Institutional Securities, Distributed Systems"),
        ("JP Morgan Chase & Co.", "3 Students", "₹90,000 / month", "Asset Management, Algorithmic Risk Systems"),
        ("Zeta Suite", "2 Students", "₹80,000 / month", "Banking-as-a-Service, High-Throughput APIs"),
        ("BrowserStack", "1 Student", "₹75,000 / month", "Cloud Infrastructure, Distributed Virtualization")
    ]

    data_table(doc, ["Host Organization", "Students Interned", "Monthly Stipend", "Technical Domain"], intern_summary, col_widths=[2.4, 1.2, 1.4, 2.0], compact=True)

    body(doc, "Global Higher Studies Progression (Batch 2025–26):", bold=True, color=NAVY)
    body(doc,
        "A significant contingent of our graduates secured admissions to premier global universities for Master of Science (MS) and doctoral degrees in Computer Science, "
        "Artificial Intelligence, Data Science, and Financial Engineering:")

    higher_studies = [
        ("Carnegie Mellon University (CMU)", "USA", "MS in Computational Data Science / Software Engineering"),
        ("Columbia University in the City of New York", "USA", "MS in Computer Science (Machine Learning Track)"),
        ("Georgia Institute of Technology", "USA", "MS in Computer Science (Cybersecurity & Systems)"),
        ("University of Southern California (USC)", "USA", "MS in Computer Science (Artificial Intelligence)"),
        ("New York University (NYU Courant)", "USA", "MS in Computer Science"),
        ("Technical University of Munich (TUM)", "Germany", "M.Sc. in Informatics"),
        ("Indian Institute of Technology (IIT) Bombay", "India", "M.Tech / Dual-Degree in Computer Science"),
        ("Indian Institute of Management (IIM) Ahmedabad / Bangalore", "India", "Post Graduate Program in Management (PGP)")
    ]

    data_table(doc, ["Admitting University / Institution", "Country", "Program of Study"], higher_studies, col_widths=[2.8, 1.0, 3.2], compact=True)

    doc.add_page_break()
