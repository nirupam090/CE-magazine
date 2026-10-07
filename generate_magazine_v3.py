"""
Master Script: Generate Version 3 of the CE Department Annual Magazine (2025–26)
Combines all 4 parts into CE_Annual_Magazine_2025-26_v3.docx
"""

import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from v3_core import create_document, OUT
from v3_part1 import build_cover_and_frontmatter, build_section_01, build_section_02, build_section_03_and_04
from v3_part2 import build_section_05, build_section_06, build_section_07
from v3_part3 import build_section_08, build_section_09_and_10, build_section_11, build_section_12, build_section_13
from v3_part4 import build_section_14, build_section_15, build_section_16, build_section_17, build_section_18, build_section_19

def main():
    print("=================================================================")
    print("  Generating CE Department Annual Magazine 2025–26 (Version 3)   ")
    print("=================================================================")
    
    doc = create_document()
    
    print("-> Building Cover Page and Hierarchical Table of Contents...")
    build_cover_and_frontmatter(doc)
    
    print("-> Building Section 01: From the Principal's Desk...")
    build_section_01(doc)
    
    print("-> Building Section 02: From the HOD's Desk...")
    build_section_02(doc)
    
    print("-> Building Sections 03 & 04: Vision, Mission, PEOs, POs, PSOs & Attainment...")
    build_section_03_and_04(doc)
    
    print("-> Building Section 05: About the Department & Comprehensive Curriculum Architecture...")
    build_section_05(doc)
    
    print("-> Building Section 06: Laboratory Facilities, Smart Classrooms & Major Equipment Added...")
    build_section_06(doc)
    
    print("-> Building Section 07: Academic Programs, Placements, Internships & Higher Studies...")
    build_section_07(doc)
    
    print("-> Building Section 08: Faculty and Staff Information...")
    build_section_08(doc)
    
    print("-> Building Sections 09 & 10: Faculty Responsibilities & Departmental Governance...")
    build_section_09_and_10(doc)
    
    print("-> Building Section 11: Consultancy, Grants, Patents & Research Publications...")
    build_section_11(doc)
    
    print("-> Building Section 12: Faculty and Students Achievements...")
    build_section_12(doc)
    
    print("-> Building Section 13: Undergraduate Projects, M.Tech Dissertations & Ph.D. Research...")
    build_section_13(doc)
    
    print("-> Building Section 14: Industry–Institute Interaction & Professional Collaborations...")
    build_section_14(doc)
    
    print("-> Building Section 15: Distinguished Alumni & Student Entrepreneurs...")
    build_section_15(doc)
    
    print("-> Building Section 16: Departmental Activities by Faculty, Students & Associations...")
    build_section_16(doc)
    
    print("-> Building Section 17: Student Support Systems, Remedial Teaching & Pedagogy...")
    build_section_17(doc)
    
    print("-> Building Section 18: SWOC Analysis Matrix & Strategic Roadmap 2026–2030...")
    build_section_18(doc)
    
    print("-> Building Section 19: Meet the Team / Editorial Board & Acknowledgments...")
    build_section_19(doc)
    
    print(f"-> Saving document to: {OUT}")
    doc.save(OUT)
    
    file_size_mb = os.path.getsize(OUT) / (1024 * 1024)
    print(f"SUCCESS: Version 3 DOCX generated successfully!")
    print(f"Output File: {OUT}")
    print(f"File Size: {file_size_mb:.2f} MB")
    print(f"Paragraphs: {len(doc.paragraphs)}")
    print(f"Tables: {len(doc.tables)}")
    print("=================================================================")

if __name__ == '__main__':
    main()
