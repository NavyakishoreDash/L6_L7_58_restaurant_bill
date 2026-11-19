import os
import zipfile
import shutil

# Staging directory
staging_dir = "LAB7_Submission_Staging"
if os.path.exists(staging_dir):
    shutil.rmtree(staging_dir)
os.makedirs(staging_dir, exist_ok=True)
os.makedirs(os.path.join(staging_dir, "screenshots"), exist_ok=True)

# 1. Copy changed code files
shutil.copy2("LAB7/restaurant_bill_parser.py", os.path.join(staging_dir, "restaurant_bill_parser.py"))
shutil.copy2("LAB7/test_restaurant_bill_parser.py", os.path.join(staging_dir, "test_restaurant_bill_parser.py"))

# 2. Copy reports (both PDF and DOCX)
shutil.copy2("LAB7/PES2UG24CS307_NavyakishoreDash_LAB7.pdf", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB7.pdf"))
shutil.copy2("LAB7/PES2UG24CS307_NavyakishoreDash_LAB7.docx", os.path.join(staging_dir, "PES2UG24CS307_NavyakishoreDash_LAB7.docx"))

# 3. Copy screenshots
for ss in ["SS1_buggy_run.png", "SS2_hypothesis_failing.png", "SS3_fixed_run.png", "SS4_hypothesis_passing.png"]:
    src = os.path.join("LAB7/screenshots", ss)
    if os.path.exists(src):
        shutil.copy2(src, os.path.join(staging_dir, "screenshots", ss))

# 4. Create github_link.txt
github_info = """========================================================================
LAB 7: FUZZING FOR INPUT VALIDATION BUGS - SUBMISSION INFORMATION
========================================================================

STUDENT DETAILS:
Name: Navyakishore Dash
SRN: PES2UG24CS307
Course: UE23CS301 - Software Engineering
Department: Department of Computer Science and Engineering
Institution: PES University

GITHUB REPOSITORY LINKS:
Primary Project Repository:
https://github.com/NavyakishoreDash/L6_L7_58_restaurant_bill
(Fork of: https://github.com/SETAPESU26/L6_L7_58_restaurant_bill)

========================================================================
SUMMARY OF CHANGED CODES & LAB 7 DELIVERABLES
========================================================================

1. restaurant_bill_parser.py (Changed & Fixed Code):
   - clean_dish_name(name): Hardened with isinstance(name, str) validation to prevent AttributeError on None or non-string inputs. Safely strips markers (*, #, ~) and whitespace before title-casing.
   - parse_item_prices(text): Hardened with type checking and safe token processing using try-except (ValueError, TypeError), gracefully handling empty strings, non-numeric tokens, and trailing delimiters while parsing valid integers.
   - split_amount(text): Hardened against missing/extra slashes by validating len(parts) == 2, explicitly checking people > 0 to prevent ZeroDivisionError, and catching (ValueError, TypeError, ZeroDivisionError).

2. test_restaurant_bill_parser.py (Completed Hypothesis Fuzz Test Suite):
   - Implemented @given(st.text() | st.none()) property-based tests for all three functions.
   - Used try-except pattern with assert False to catch and report unexpected crashes.

3. Reports & Documentation:
   - PES2UG24CS307_NavyakishoreDash_LAB7.pdf: Complete PDF report with SS1-SS4 screenshots and reflection answers.
   - PES2UG24CS307_NavyakishoreDash_LAB7.docx: Formatted Word document with all test cases, bugs identified, and fixes.
   - screenshots/: High-resolution terminal capture images (SS1 to SS4).
"""

with open(os.path.join(staging_dir, "github_link.txt"), "w", encoding="utf-8") as f:
    f.write(github_info)

with open(os.path.join(staging_dir, "README.txt"), "w", encoding="utf-8") as f:
    f.write(github_info)

# 5. Create zip archives
zip_filenames = [
    "PES2UG24CS307_Lab7.zip",
    "PES2UG24CS307.zip",
    "PES2UG24CS307_NavyakishoreDash_LAB7.zip"
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
print("Packaging complete!")
