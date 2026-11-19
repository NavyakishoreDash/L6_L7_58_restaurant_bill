import os
import zipfile
import shutil

# Staging directory for Lab 6
staging_dir = "LAB6_Submission_Staging"
if os.path.exists(staging_dir):
    shutil.rmtree(staging_dir)
os.makedirs(staging_dir, exist_ok=True)
os.makedirs(os.path.join(staging_dir, "screenshots"), exist_ok=True)

# 1. Copy main python code and test code
shutil.copy2("LAB6/restaurant_bill.py", os.path.join(staging_dir, "restaurant_bill.py"))
shutil.copy2("LAB6/test_restaurant_bill.py", os.path.join(staging_dir, "test_restaurant_bill.py"))

# 2. Copy Word doc and PDF reports
shutil.copy2("LAB6/PES2UG24CS307_NavyakishoreDash_LAB6.docx", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB6.docx"))
shutil.copy2("LAB6/PES2UG24CS307_NavyakishoreDash_LAB6.pdf", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB6.pdf"))

# 3. Copy screenshot
shutil.copy2("LAB6/screenshots/SS_lab6_coverage.png", os.path.join(staging_dir, "screenshots", "SS_lab6_coverage.png"))

# 4. Create github_link.txt
github_info = """========================================================================
LAB 6: CODE AND BRANCH COVERAGE - SUBMISSION INFORMATION
========================================================================

STUDENT DETAILS:
Name: Navyakishore Dash
SRN: PES2UG24CS307
Course: UE23CS301 - Software Engineering
Department: Department of Computer Science and Engineering
Institution: PES University

GITHUB REPOSITORY LINK (Copy-paste this into the form):
https://github.com/NavyakishoreDash/L6_L7_58_restaurant_bill
(Base Upstream: https://github.com/SETAPESU26/L6_L7_58_restaurant_bill)

========================================================================
SUMMARY OF LAB 6 DELIVERABLES & COVERAGE ACHIEVED
========================================================================

1. Target Module: restaurant_bill.py
   - bill_total(section, food_total, guests, is_happy_hour)
   - split_bill(total, people, tip_percent)
   - table_wait(party_size, free_tables, has_reservation)

2. Test Suite: test_restaurant_bill.py
   - 21 comprehensive test cases exercising all decision branches and boundary conditions.
   - Handled all exception cases (ValueError for unknown section, invalid orders, invalid people, negative tips).

3. Coverage Results (pytest-cov):
   ========================================================================
   Name                 Stmts   Miss Branch BrPart  Cover   Missing
   ------------------------------------------------------------------------
   restaurant_bill.py      33      0     22      0   100%
   ------------------------------------------------------------------------
   TOTAL                   33      0     22      0   100%
   ========================================================================
   - Statement Coverage: 100% (33/33 statements)
   - Branch Coverage: 100% (22/22 branches)
   - Test Results: 21 passed in 1.35s

4. Documentation:
   - PES2UG24CS307_NavyakishoreDash_LAB6.docx: Word document with screenshots, main python code, test code, and decision table.
   - PES2UG24CS307_NavyakishoreDash_LAB6.pdf: PDF document version.
   - screenshots/SS_lab6_coverage.png: Terminal execution screenshot.
"""

with open(os.path.join(staging_dir, "github_link.txt"), "w", encoding="utf-8") as f:
    f.write(github_info)

with open(os.path.join(staging_dir, "README.txt"), "w", encoding="utf-8") as f:
    f.write(github_info)

# 5. Create zip archives
zip_filenames = [
    "PES2UG24CS307_Lab6.zip",
    "PES2UG24CS307_NavyakishoreDash_LAB6.zip"
]

for zname in zip_filenames:
    with zipfile.ZipFile(zname, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(staging_dir):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, staging_dir)
                zipf.write(full_path, rel_path)
    size_kb = os.path.getsize(zname) / 1024
    print(f"Created {zname} ({size_kb:.1f} KB)")

# Cleanup staging
shutil.rmtree(staging_dir)
print("Lab 6 packaging complete!")
