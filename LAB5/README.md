# Lab 5: Static Code Analysis — Support Ticket Tracker

**Student:** Navyakishore Dash  
**SRN:** PES2UG24CS307  
**Section:** Section -E  
**Assigned File:** `support_ticket_tracker.py`  
**Assigned Repository:** `https://github.com/SETAPESU26/L5_58_support-ticket-tracker`  
**Student Fork:** `https://github.com/NavyakishoreDash/L5_58_support-ticket-tracker`  

---

## About the Program

A customer support ticket tracker. It creates tickets with a priority and category, assigns them to agents, updates their status, searches and reports on them, and saves them to a JSON file.

The original code worked, but had significant quality, style, and security problems of the kind Pylint, Flake8, and Bandit are built to find.

## Analysis & Remediation Summary

1. **Before Static Analysis:**
   - **Pylint Score:** `5.00/10` (33 messages: dangerous mutable default argument, bare except, file handles left open, global variables, missing docstrings, long lines, unused imports/variables)
   - **Bandit Security Scan:** `2 issues` (Medium: `B307` eval used; Low: `B110` try_except_pass)
   - **Flake8 Linter:** `12 errors` (unused imports `sys` and `random`, lines exceeding 79 chars, bare except, comparison to None, unused variable, trailing whitespace)

2. **After Refactoring:**
   - **Pylint Score:** `10.00/10` (0 messages)
   - **Bandit Security Scan:** `0 issues` (No issues identified)
   - **Flake8 Linter:** `0 errors` (Clean exit)

## Deliverables in this Submission

1. `support_ticket_tracker.py`: Cleaned, robust, and PEP 8-compliant version (10.00/10 Pylint score, 0 Bandit issues, 0 Flake8 errors).
2. `issues.md`: Comprehensive table documenting all 10 identified and resolved issues with line numbers, descriptions, and fix strategies.
3. `reflection.md`: Detailed answers to all four lab reflection questions.
4. `reports/before/`: Raw output logs from Pylint, Bandit, and Flake8 before refactoring.
5. `reports/after/`: Raw output logs from Pylint, Bandit, and Flake8 after refactoring.
6. `PES2UG24CS307_NavyakishoreDash_LAB5_Report.pdf`: Formal 4-page academic lab report.
7. `PES2UG24CS307_NavyakishoreDash_LAB5_Report.docx`: Formatted Word document version.
