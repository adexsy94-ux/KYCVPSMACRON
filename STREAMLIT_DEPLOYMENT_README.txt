KYC/VPS/MACRON RECONCILIATION — NO-OPENPYXL DROP-IN FIX
========================================================

WHY THIS VERSION EXISTS
-----------------------
The deployed app repeatedly failed because Streamlit Cloud did not have openpyxl.
This version removes openpyxl from XLSX reading entirely. It reads .xlsx files
with Python's standard ZIP/XML libraries, while keeping the reconciliation logic,
KYC mapping, management report, Cash Posting, and Excel export.

DROP-IN FILE NAME
-----------------
The Python file is deliberately named:
    kyc_driven_vps_macron_reconciliation_v12.py

so you can replace your existing Streamlit entrypoint without changing the app
settings.

REPOSITORY ROOT SHOULD CONTAIN
------------------------------
    kyc_driven_vps_macron_reconciliation_v12.py
    requirements.txt

requirements.txt MUST be a real file and MUST contain one dependency per line.
Do not paste all requirements onto one line.

IMPORTANT STREAMLIT DEPENDENCY RULE
-----------------------------------
Streamlit Community Cloud uses the first dependency file it finds. If your repo
contains uv.lock, Pipfile, environment.yml, or another higher-priority dependency
file beside the app, requirements.txt may be ignored. Keep only the dependency
file you actually intend to use.

WHAT WAS TESTED
---------------
Tested with the supplied sample files while imports of openpyxl were deliberately
blocked.

Source coverage:
  Bank transactions: 354 / 354, missing 0, duplicate use 0
  VPS:               345 / 345, missing 0, duplicate use 0
  Paymeter:          340 / 340, missing 0, duplicate use 0
  Macron:            312 / 312, missing 0, duplicate use 0

Cash Posting:
  358 rows
  exactly 7 columns:
    Transaction Date
    Customer Name
    Account Number
    Cash In
    Cash Out
    Interest
    Source
  missing Transaction Date: 0

The generated totals matched the previous V12 test output:
  Cash In:    NGN 11,144,783.00
  Cash Out:   NGN 18,425,223.52
  Interest:   NGN 21,364.73

DATE HANDLING
-------------
Paymeter uses month/day order, e.g. 10/11/2026 = 11-Oct-2026.
Macron and Bank use day/month order.
VPS ISO/UTC timestamps are converted to Lagos local time.
Raw Excel serial dates are also supported without openpyxl.

DEPLOYMENT
----------
1. Replace your existing Python file with the supplied file.
2. Put requirements.txt in the same folder as the app or repository root.
3. Commit both files to GitHub.
4. Check the repo for uv.lock, Pipfile, or environment.yml. Remove/rename an
   unintended competing dependency file.
5. In Streamlit Cloud, reboot the app.

openpyxl is intentionally NOT required by this version.
