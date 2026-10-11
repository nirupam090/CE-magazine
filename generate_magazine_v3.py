"""
build_master_magazine.py - Complete Master Builder for CE Annual Magazine AY 2025-26
"""

import os, sys, io

# Set working directory to scratch
SCRATCH = r'C:\Users\Nirupam\.gemini\antigravity-ide\brain\d84a1836-479b-4390-be26-b801a6bee109\scratch'
sys.path.insert(0, SCRATCH)

from magazine_core import create_document
from verified_builder_part1 import build_cover_and_frontmatter, build_section_01, build_section_02, build_section_03_and_04
from verified_builder_part2 import build_section_05, build_section_06, build_section_07
from verified_builder_part3 import build_section_08, build_section_09_and_10, build_section_11, build_section_12, build_section_13
from verified_builder_part4 import build_section_14, build_section_15, build_section_16, build_section_17, build_section_18, build_section_19

TARGET_PATH = r'C:\Users\Nirupam\Documents\magazine ce\CE_Annual_Magazine_2025-26_v3.docx'

def main():
    print("=" * 70)
    print("GENERATING VERIFIED CE DEPARTMENT ANNUAL MAGAZINE AY 2025-26")
    print("=" * 70)

    doc = create_document()

    print("-> 1. Cover Page & Front Matter...")
    build_cover_and_frontmatter(doc)

    print("-> 2. Section 01: From the Principal's Desk (Verbatim from CSE Magazine)...")
    build_section_01(doc)

    print("-> 3. Section 02: From the Head of Department's Desk...")
    build_section_02(doc)

    print("-> 4. Sections 03 & 04: Vision, Mission, PEOs, POs, PSOs & Audited Attainment...")
    build_section_03_and_04(doc)

    print("-> 5. Section 05: About the Department & Comprehensive Curriculum Architecture...")
    build_section_05(doc)

    print("-> 6. Section 06: Laboratory Facilities, Smart Classrooms & Major Equipment Added...")
    build_section_06(doc)

    print("-> 7. Section 07: Academic Programs, Placements, Internships & Higher Studies...")
    build_section_07(doc)

    print("-> 8. Section 08: Faculty and Staff Information...")
    build_section_08(doc)

    print("-> 9. Sections 09 & 10: Faculty Responsibilities & Departmental Governance...")
    build_section_09_and_10(doc)

    print("-> 10. Section 11: Consultancy, Grants, Patents & Research Publications...")
    build_section_11(doc)

    print("-> 11. Section 12: Faculty and Student Achievements...")
    build_section_12(doc)

    print("-> 12. Section 13: Undergraduate Projects, M.Tech Dissertations & Ph.D. Research...")
    build_section_13(doc)

    print("-> 13. Section 14: Industry–Institute Interaction...")
    build_section_14(doc)

    print("-> 14. Section 15: Distinguished Alumni & Mentorship...")
    build_section_15(doc)

    print("-> 15. Section 16: Departmental Activities & Student Associations...")
    build_section_16(doc)

    print("-> 16. Section 17: Student Support Systems & Pedagogy...")
    build_section_17(doc)

    print("-> 17. Section 18: SWOC Analysis Matrix & Departmental Goals...")
    build_section_18(doc)

    print("-> 18. Section 19: Meet the Editorial Board & Closing...")
    build_section_19(doc)

    print(f"\n-> Saving document to: {TARGET_PATH}")
    doc.save(TARGET_PATH)

    file_size_mb = os.path.getsize(TARGET_PATH) / (1024 * 1024)
    print("\n" + "=" * 70)
    print("SUCCESS: VERIFIED MAGAZINE GENERATED SUCCESSFULLY!")
    print(f"Output File:     {TARGET_PATH}")
    print(f"File Size:       {file_size_mb:.2f} MB")
    print(f"Total Paragraphs: {len(doc.paragraphs)}")
    print(f"Total Tables:    {len(doc.tables)}")
    print("=" * 70)

if __name__ == '__main__':
    main()
