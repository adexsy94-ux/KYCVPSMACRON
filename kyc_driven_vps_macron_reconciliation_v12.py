# kyc_driven_vps_macron_reconciliation_v12.py
# ================================================================
# KYC-DRIVEN VPS / MACRON RECONCILIATION
# ================================================================
# Required business sequence:
#   0. Clean/read all five sources with source-specific date handling.
#   1. Build supporting Paymeter <-> Macron history by exact reference.
#   2. Match VPS <-> Bank credits; retain unrelated dated Bank credits/debits.
#   3. Build the master identity table from the uploaded KYC file.
#   4. Attach KYC to Paymeter-Macron and VPS-Bank histories.
#   5. Bridge VPS <-> Paymeter by account, gross value and time.
#   6. Reconcile VPS directly against Macron using reference, KYC, value and time.
#   7. Compare Expected Macron Token = VPS transaction_amount_minor - N100.
#   8. Build Customer Summary, hard source tie-outs, management reports and Cash Posting.
# ================================================================

import csv
import io
import re
from collections import defaultdict, deque
from pathlib import Path

import numpy as np
import pandas as pd

try:
    import streamlit as st
except Exception:
    st = None


AMOUNT_TOLERANCE = 0.01
DATE_TOLERANCE_DAYS = 2
TRANSACTION_TIME_TOLERANCE_HOURS = 36
HIGH_CONFIDENCE_MINUTES = 30
MACRON_TOKEN_CHARGE = 100.00

# ================================================================
# OPTION 2 MANAGEMENT REPORT SETTINGS
# ================================================================
# Exposure classification is based on EACH customer's individual
# share of the relevant total outstanding (recovery or token due).
OPTION2_CRITICAL_SHARE = 0.10       # >= 10%
OPTION2_HIGH_SHARE = 0.02           # 2% to <10%
OPTION2_SIGNIFICANT_SHARE = 0.005   # 0.5% to <2%

# Purchase-frequency bands are deliberately kept separate from
# financial materiality. They describe customer activity only.
OPTION2_VERY_FREQUENT_MIN = 13
OPTION2_FREQUENT_MIN = 7
OPTION2_REGULAR_MIN = 4

PAYMETER_HEADER = [
    "Transaction ID", "Created At", "Updated Date", "Customer Name",
    "Phone Number", "Bank Name", "Disco Name", "District Name",
    "Account Number", "Disco Account Number", "System Fee Per Disco",
    "Meter Number", "Address", "Transaction Amount", "Input Amount",
    "System Charge", "Disco Commission Type", "Disco Amount",
    "Disco System Commission Fee", "Disco Commission Fee Value",
    "Disco System Commission Cap Fee", "RRN", "Reference", "User Type",
    "Earning Partner ID", "Earning Partner Name", "Earning Partner Fee",
    "Earning Partner Commission Type", "Earning Fee Value",
    "Earning Partner Cap Fee", "District Manager Fee",
    "Agent Commission Type", "Agent Commission Fee",
    "Agent Commission Percentage Type Fee Value", "Agent Commission Cap Fee",
    "Bank Commission Type", "Bank Commission Fee", "Bank Commission Fee Value",
    "Bank Commission Fee Cap", "Bank Tax Fee", "Fee Due To System",
    "Message ID", "SMS Delivery Status", "Status Checked",
]


def clean_id(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    s = str(value).strip()
    if s.lower() in {"nan", "none", "nat"}:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def normalize_meter(value):
    s = clean_id(value)
    if not s:
        return ""
    return (s.lstrip("0") or "0") if s.isdigit() else s.upper()


def to_number(value):
    s = clean_id(value).replace(",", "")
    if not s:
        return np.nan
    s = re.sub(r"[^0-9.\-]", "", s)
    return pd.to_numeric(s, errors="coerce")


def numeric_series(series):
    s = pd.Series(series, dtype="string")
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"[^0-9.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def amount_equal(a, b, tolerance=AMOUNT_TOLERANCE):
    try:
        return pd.notna(a) and pd.notna(b) and abs(float(a) - float(b)) <= tolerance
    except Exception:
        return False


def parse_datetime(series, utc=False, dayfirst=True):
    if utc:
        d = pd.to_datetime(series, errors="coerce", utc=True)
        try:
            return d.dt.tz_convert("Africa/Lagos").dt.tz_localize(None)
        except Exception:
            return d.dt.tz_localize(None)
    return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)


def unique_join(values):
    out, seen = [], set()
    for v in values:
        s = clean_id(v)
        if s and s not in seen:
            seen.add(s)
            out.append(s)
    return " | ".join(out)


def most_common_text(values):
    counts = defaultdict(int)
    order = []
    for v in values:
        s = clean_id(v)
        if not s:
            continue
        if s not in counts:
            order.append(s)
        counts[s] += 1
    if not counts:
        return ""
    return max(order, key=lambda x: counts[x])


def split_multi(value):
    return [x.strip() for x in clean_id(value).split("|") if x.strip()]


def prefixed_record(row, prefix):
    return {f"{prefix}{k}": v for k, v in row.items()}


def read_bytes(file_obj):
    if isinstance(file_obj, (str, Path)):
        return Path(file_obj).read_bytes()
    if isinstance(file_obj, (bytes, bytearray)):
        return bytes(file_obj)
    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()
    try:
        file_obj.seek(0)
    except Exception:
        pass
    return file_obj.read()


def excel_input(file_obj):
    if isinstance(file_obj, (str, Path)):
        return file_obj
    return io.BytesIO(read_bytes(file_obj))


def suffix_of(file_obj):
    if isinstance(file_obj, (str, Path)):
        return Path(file_obj).suffix.lower()
    return Path(getattr(file_obj, "name", "")).suffix.lower()


def require_columns(df, columns, source):
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"{source} missing required column(s): {', '.join(missing)}")


def progress_call(callback, pct, message):
    if callback:
        callback(pct, message)


# ================================================================
# 0. CLEAN PAYMETER BEFORE ANY MERGE
# ================================================================

def _blank_header(value):
    s = clean_id(value)
    return not s or bool(re.fullmatch(r"Unnamed:\s*\d+", s, flags=re.I))


def _normalise_paymeter_header(raw_header):
    header = [str(x).replace("\ufeff", "").strip() for x in raw_header]
    ignored = 0
    while header and _blank_header(header[-1]):
        header.pop()
        ignored += 1
    if header != PAYMETER_HEADER:
        missing = [c for c in PAYMETER_HEADER if c not in header]
        if missing:
            raise ValueError("Paymeter missing required column(s): " + ", ".join(missing))
        diffs = []
        for i, expected in enumerate(PAYMETER_HEADER):
            actual = header[i] if i < len(header) else "<missing>"
            if actual != expected:
                diffs.append(f"column {i+1}: expected '{expected}', found '{actual}'")
            if len(diffs) >= 5:
                break
        raise ValueError(
            "Paymeter named business columns differ from the expected 44-column layout. "
            + "; ".join(diffs)
        )
    return header, ignored


def _find_true_pm_amount_start(row):
    """Find the true Transaction Amount position after an Address spill."""
    tx_idx = PAYMETER_HEADER.index("Transaction Amount")
    rrn_offset = PAYMETER_HEADER.index("RRN") - tx_idx
    ref_offset = PAYMETER_HEADER.index("Reference") - tx_idx
    user_offset = PAYMETER_HEADER.index("User Type") - tx_idx
    addr_idx = PAYMETER_HEADER.index("Address")
    candidates = []

    for pos in range(addr_idx + 1, max(addr_idx + 2, len(row) - 2)):
        if pos + 2 >= len(row):
            break
        tx, inp, charge = to_number(row[pos]), to_number(row[pos + 1]), to_number(row[pos + 2])
        if pd.isna(tx) or pd.isna(inp) or pd.isna(charge):
            continue
        score = 5
        if amount_equal(float(inp) - float(charge), tx):
            score += 10
        if pos + rrn_offset < len(row) and clean_id(row[pos + rrn_offset]):
            score += 2
        if pos + ref_offset < len(row) and clean_id(row[pos + ref_offset]):
            score += 2
        if pos + user_offset < len(row) and clean_id(row[pos + user_offset]):
            score += 1
        candidates.append((score, pos))

    if not candidates:
        return None, 0
    candidates.sort(key=lambda x: (-x[0], x[1]))
    return candidates[0][1], candidates[0][0]


def _repair_paymeter_row(raw_row):
    expected = len(PAYMETER_HEADER)
    addr_idx = PAYMETER_HEADER.index("Address")
    tx_idx = PAYMETER_HEADER.index("Transaction Amount")

    # A perfect-width row can be used directly if the amount triplet is valid.
    if len(raw_row) == expected:
        tx, inp, chg = map(to_number, raw_row[tx_idx:tx_idx + 3])
        if pd.notna(tx) and pd.notna(inp) and pd.notna(chg):
            return list(raw_row), "CLEAN", 0, "HIGH"

    true_pos, score = _find_true_pm_amount_start(raw_row)
    if true_pos is not None:
        first_part = list(raw_row[:addr_idx])
        original_address = raw_row[addr_idx] if addr_idx < len(raw_row) else ""
        tail = list(raw_row[true_pos:])
        tail_needed = expected - tx_idx
        tail = (tail + [""] * tail_needed)[:tail_needed]
        rebuilt = first_part + [original_address] + tail
        rebuilt = (rebuilt + [""] * expected)[:expected]
        deleted = max(true_pos - (addr_idx + 1), 0)
        status = "REPAIRED - ADDRESS SPILL DELETED" if deleted else "CLEAN - EXPORT PADDING NORMALISED"
        return rebuilt, status, deleted, "HIGH" if score >= 15 else "MEDIUM"

    # Last-resort preservation. Keep first/original address and retain row.
    left = list(raw_row[:addr_idx + 1])
    rest = list(raw_row[addr_idx + 1:])
    rebuilt = (left + rest + [""] * expected)[:expected]
    return rebuilt, "REVIEW REQUIRED", max(len(raw_row) - expected, 0), "LOW"


def clean_paymeter(file_obj):
    if suffix_of(file_obj) in {".xlsx", ".xls"}:
        df = pd.read_excel(excel_input(file_obj), dtype=str)
        require_columns(df, PAYMETER_HEADER, "Paymeter")
        df = df[PAYMETER_HEADER].copy()
        df["PM_Cleaning_Status"] = "CLEAN XLS/XLSX"
        df["PM_Address_Spill_Cells_Deleted"] = 0
        df["PM_Cleaning_Confidence"] = "HIGH"
        df["PM_Source_Line"] = np.arange(2, len(df) + 2)
        ignored_header = 0
    else:
        rows = list(csv.reader(io.StringIO(read_bytes(file_obj).decode("utf-8-sig", errors="replace"))))
        if not rows:
            raise ValueError("Paymeter report is empty.")
        _, ignored_header = _normalise_paymeter_header(rows[0])
        records = []
        for line, raw_row in enumerate(rows[1:], start=2):
            if not any(clean_id(x) for x in raw_row):
                continue
            rebuilt, status, deleted, confidence = _repair_paymeter_row(raw_row)
            rec = dict(zip(PAYMETER_HEADER, rebuilt))
            rec["PM_Cleaning_Status"] = status
            rec["PM_Address_Spill_Cells_Deleted"] = deleted
            rec["PM_Cleaning_Confidence"] = confidence
            rec["PM_Source_Line"] = line
            records.append(rec)
        df = pd.DataFrame(records)

    for c in [
        "Transaction Amount", "Input Amount", "System Charge", "System Fee Per Disco",
        "Disco Amount", "Disco System Commission Fee", "Disco Commission Fee Value",
        "Earning Partner Fee", "Bank Commission Fee", "Bank Commission Fee Value",
        "Bank Tax Fee", "Fee Due To System",
    ]:
        if c in df.columns:
            df[c] = numeric_series(df[c])
    for c in ["Transaction ID", "Account Number", "Disco Account Number", "Meter Number", "RRN", "Reference"]:
        df[c] = df[c].map(clean_id)

    def _id_export_quality(value):
        s = clean_id(value)
        if not s:
            return "BLANK"
        if re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)[Ee][+-]?\d+", s):
            return "EXPORT PRECISION LOST"
        return "OK"

    df["PM_RRN_Quality"] = df["RRN"].map(_id_export_quality)
    df["PM_Meter_Quality"] = df["Meter Number"].map(_id_export_quality)
    # Paymeter exports dates in US month/day order (for example 10/11/2026 = 11-Oct-2026).
    # Keep this source-specific; using the shared day-first parser would silently swap ambiguous dates.
    df["PM_DateTime"] = pd.to_datetime(df["Created At"], errors="coerce", format="mixed", dayfirst=False)
    df["PM_Date"] = df["PM_DateTime"].dt.normalize()
    df["PM_Row"] = np.arange(len(df))
    df.attrs["ignored_blank_header_columns"] = ignored_header
    df.attrs["repaired_rows"] = int((pd.to_numeric(df["PM_Address_Spill_Cells_Deleted"], errors="coerce").fillna(0) > 0).sum())
    df.attrs["spill_cells_deleted"] = int(pd.to_numeric(df["PM_Address_Spill_Cells_Deleted"], errors="coerce").fillna(0).sum())
    return df


# ================================================================
# SOURCE READERS
# ================================================================

def read_vps(file_obj):
    df = pd.read_excel(excel_input(file_obj), dtype=str)
    require_columns(df, [
        "session_id", "settlement_ref", "transaction_amount_minor", "settled_amount_minor",
        "charge_amount_minor", "source_acct_name", "source_acct_no", "virtual_acct_no", "created_at",
    ], "VPS")
    for c in ["transaction_amount_minor", "settled_amount_minor", "charge_amount_minor", "vat_amount_minor"]:
        if c in df.columns:
            df[c] = numeric_series(df[c])
    for c in ["session_id", "settlement_ref", "virtual_acct_no", "source_acct_no"]:
        if c in df.columns:
            df[c] = df[c].map(clean_id)
    df["VPS_DateTime"] = parse_datetime(df["created_at"], utc=True)
    df["VPS_Date"] = df["VPS_DateTime"].dt.normalize()
    df["VPS_Row"] = np.arange(len(df))
    return df


def read_macron(file_obj):
    df = pd.read_excel(excel_input(file_obj), dtype=str)
    require_columns(df, ["DATE & TIME", "TRANSACTION ID", "METER NUMBER", "ACCOUNT NUMBER", "AMOUNT", "STATUS", "REFERENCE ID"], "Macron")
    df["AMOUNT"] = numeric_series(df["AMOUNT"])
    for c in ["TRANSACTION ID", "METER NUMBER", "ACCOUNT NUMBER", "REFERENCE ID"]:
        df[c] = df[c].map(clean_id)
    # Macron exports dates in day/month order (for example 11/10/2026 = 11-Oct-2026).
    df["MAC_DateTime"] = pd.to_datetime(df["DATE & TIME"], errors="coerce", format="mixed", dayfirst=True)
    df["MAC_Date"] = df["MAC_DateTime"].dt.normalize()
    df["MAC_Row"] = np.arange(len(df))
    return df


def read_bank(file_obj):
    raw = pd.read_excel(excel_input(file_obj), header=None, dtype=str)
    header_row = None
    for i in range(min(100, len(raw))):
        vals = [clean_id(x) for x in raw.iloc[i].tolist()]
        if "Narration" in vals and "Credit" in vals:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not locate Bank transaction header containing Narration and Credit.")
    headers = [clean_id(x) or f"Unnamed_{j}" for j, x in enumerate(raw.iloc[header_row].tolist())]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)
    require_columns(df, ["Narration", "Credit"], "Bank")
    for c in ["Debit", "Credit", "Current Balance"]:
        if c in df.columns:
            df[c] = numeric_series(df[c])
    date_col = "Actual Transaction Date" if "Actual Transaction Date" in df.columns else "Post Date"
    # Bank statement dates are day/month. A numeric Total line appears below the
    # transactions in the sample; it must never be treated as a bank receipt.
    df["BANK_DateTime"] = pd.to_datetime(df[date_col], errors="coerce", format="mixed", dayfirst=True)
    df["BANK_Date"] = df["BANK_DateTime"].dt.normalize()
    has_amount = pd.Series(False, index=df.index)
    if "Debit" in df.columns:
        has_amount = has_amount | df["Debit"].notna()
    if "Credit" in df.columns:
        has_amount = has_amount | df["Credit"].notna()
    df = df[df["BANK_DateTime"].notna() & has_amount].copy().reset_index(drop=True)
    df["BANK_Row"] = np.arange(len(df))
    return df


def read_kyc(file_obj):
    df = pd.read_excel(excel_input(file_obj), dtype=str)
    require_columns(df, ["PAYMETER - Account Number", "PAYMETER - Customer Name", "MACRON ACCOUNT NUMBER"], "KYC")
    if "METER NUMBER" not in df.columns:
        df["METER NUMBER"] = ""
    for c in ["PAYMETER - Account Number", "PAYMETER - Customer Name", "METER NUMBER", "MACRON ACCOUNT NUMBER"]:
        df[c] = df[c].map(clean_id)
    df["KYC_Row"] = np.arange(len(df))
    return df


# ================================================================
# GENERIC EXACT ONE-TO-ONE MATCH
# ================================================================

def exact_queue_map(left_series, right_series):
    queues = defaultdict(deque)
    for ri, v in right_series.items():
        key = clean_id(v)
        if key:
            queues[key].append(ri)
    mapping, used = {}, set()
    for li, v in left_series.items():
        key = clean_id(v)
        if key and queues.get(key):
            ri = queues[key].popleft()
            mapping[li] = ri
            used.add(ri)
    return mapping, used


# ================================================================
# 1. MERGE PAYMETER WITH MACRON FIRST
# ================================================================

def merge_paymeter_macron(pay, mac):
    mapping, used_mac = exact_queue_map(pay["Reference"], mac["REFERENCE ID"])
    pay_records = pay.to_dict("index")
    mac_records = mac.to_dict("index")
    rows = []

    for pi, p in pay_records.items():
        mi = mapping.get(pi)
        m = mac_records.get(mi) if mi is not None else None
        rec = {
            "PMMAC_Row": len(rows),
            "PMMAC_Match_Status": "MATCHED" if m is not None else "PAYMETER ONLY - NO MACRON",
            "PMMAC_Match_Method": "Paymeter Reference = Macron REFERENCE ID" if m is not None else "",
        }
        rec.update(prefixed_record(p, "PM_"))
        if m is not None:
            rec.update(prefixed_record(m, "MAC_"))
        rows.append(rec)

    for mi, m in mac_records.items():
        if mi in used_mac:
            continue
        rec = {
            "PMMAC_Row": len(rows),
            "PMMAC_Match_Status": "MACRON ONLY - NO PAYMETER",
            "PMMAC_Match_Method": "",
        }
        rec.update(prefixed_record(m, "MAC_"))
        rows.append(rec)

    out = pd.DataFrame(rows)
    # Ensure key prefixed columns exist even if a source is absent in all rows.
    for c in ["PM_RRN", "PM_Reference", "PM_Account Number", "PM_Disco Account Number", "PM_Meter Number", "PM_Customer Name", "PM_Phone Number", "PM_Input Amount", "PM_Transaction Amount", "PM_System Charge", "PM_PM_Date", "PM_PM_DateTime", "MAC_ACCOUNT NUMBER", "MAC_METER NUMBER", "MAC_AMOUNT", "MAC_REFERENCE ID", "MAC_MAC_Date", "MAC_MAC_DateTime"]:
        if c not in out.columns:
            out[c] = np.nan
    return out


# ================================================================
# 2. MERGE VPS WITH BANK STATEMENT
# ================================================================

def merge_vps_bank(vps, bank, date_tolerance=DATE_TOLERANCE_DAYS):
    # VPS receipts can only match bank credits, but the full dated bank
    # transaction population (credits and debits) is retained for reporting.
    credits = bank[bank["Credit"].notna()].copy()
    bank_credit_records = credits.to_dict("index")
    bank_all_records = bank.to_dict("index")
    bank_records = bank_credit_records
    vps_records = vps.to_dict("index")

    # Primary exact session ID from Bank Narration.
    token_queues = defaultdict(deque)
    for bi, b in bank_records.items():
        narration = clean_id(b.get("Narration"))
        for token in re.findall(r"\d{12,}", narration):
            token_queues[token].append(bi)

    vps_to_bank, used_bank = {}, set()
    match_method = {}
    for vi, v in vps_records.items():
        sid = clean_id(v.get("session_id"))
        q = token_queues.get(sid)
        while q and q[0] in used_bank:
            q.popleft()
        if q:
            bi = q.popleft()
            vps_to_bank[vi] = bi
            used_bank.add(bi)
            match_method[vi] = "VPS session_id found in Bank Narration"

    # Secondary unique date + settled amount.
    amount_date = defaultdict(list)
    for bi, b in bank_records.items():
        if bi in used_bank or pd.isna(b.get("Credit")) or pd.isna(b.get("BANK_Date")):
            continue
        amount_date[(pd.Timestamp(b["BANK_Date"]), round(float(b["Credit"]), 2))].append(bi)

    for vi, v in vps_records.items():
        if vi in vps_to_bank or pd.isna(v.get("settled_amount_minor")) or pd.isna(v.get("VPS_Date")):
            continue
        base_date = pd.Timestamp(v["VPS_Date"])
        amount = round(float(v["settled_amount_minor"]), 2)
        candidates = []
        for d in range(-date_tolerance, date_tolerance + 1):
            candidates.extend([x for x in amount_date.get((base_date + pd.Timedelta(days=d), amount), []) if x not in used_bank])
        candidates = list(dict.fromkeys(candidates))
        if len(candidates) == 1:
            bi = candidates[0]
            vps_to_bank[vi] = bi
            used_bank.add(bi)
            match_method[vi] = "Unique VPS settled amount + date"

    rows = []
    for vi, v in vps_records.items():
        bi = vps_to_bank.get(vi)
        b = bank_records.get(bi) if bi is not None else None
        rec = {
            "VPSBANK_Row": len(rows),
            "VPSBANK_Match_Status": "MATCHED" if b is not None else "VPS ONLY - NO BANK",
            "VPSBANK_Match_Method": match_method.get(vi, ""),
        }
        rec.update(prefixed_record(v, "VPS_"))
        if b is not None:
            rec.update(prefixed_record(b, "BANK_"))
        rows.append(rec)

    # Retain every genuine bank movement that was not consumed by a VPS-credit
    # match. This includes unrelated credits and all bank debits/outflows.
    for bi, b in bank_all_records.items():
        if bi in used_bank:
            continue
        movement = "DEBIT" if pd.notna(b.get("Debit")) else "CREDIT" if pd.notna(b.get("Credit")) else "MOVEMENT"
        rec = {
            "VPSBANK_Row": len(rows),
            "VPSBANK_Match_Status": f"BANK ONLY - {movement} - NO VPS",
            "VPSBANK_Match_Method": "",
        }
        rec.update(prefixed_record(b, "BANK_"))
        rows.append(rec)

    out = pd.DataFrame(rows)
    for c in ["VPS_virtual_acct_no", "VPS_settlement_ref", "VPS_session_id", "VPS_transaction_amount_minor", "VPS_settled_amount_minor", "VPS_charge_amount_minor", "VPS_VPS_Date", "VPS_VPS_DateTime", "BANK_Debit", "BANK_Credit", "BANK_Narration", "BANK_BANK_Date", "BANK_BANK_DateTime"]:
        if c not in out.columns:
            out[c] = np.nan
    return out


# ================================================================
# 3. BUILD MASTER KYC FROM THE UPLOADED KYC FILE
# ================================================================

def build_kyc(kyc_raw, pay, mac, vps_bank_merge):
    """Build one customer identity row per Paymeter/VPS virtual account.

    Uploaded KYC is authoritative. Paymeter/VPS/Macron observations enrich the
    master and are used to expose conflicts, never to silently overwrite KYC.
    """
    raw = kyc_raw.copy()
    pay = pay.copy()
    mac = mac.copy()
    vps_rows = vps_bank_merge[vps_bank_merge["VPS_virtual_acct_no"].map(clean_id).ne("")].copy()

    kyc_groups = {k: g.copy() for k, g in raw.groupby(raw["PAYMETER - Account Number"].map(clean_id), sort=False) if k}
    pay_groups = {k: g.copy() for k, g in pay.groupby(pay["Account Number"].map(clean_id), sort=False) if k}
    vps_groups = {k: g.copy() for k, g in vps_rows.groupby(vps_rows["VPS_virtual_acct_no"].map(clean_id), sort=False) if k}
    mac_groups = {k: g.copy() for k, g in mac.groupby(mac["ACCOUNT NUMBER"].map(clean_id), sort=False) if k}

    uploaded_mac_to_pay = defaultdict(set)
    for _, rr in raw.iterrows():
        pa = clean_id(rr.get("PAYMETER - Account Number"))
        ma = clean_id(rr.get("MACRON ACCOUNT NUMBER"))
        if pa and ma:
            uploaded_mac_to_pay[ma].add(pa)

    all_accounts = sorted(set(kyc_groups) | set(pay_groups) | set(vps_groups))
    records = []
    for acct in all_accounts:
        kg = kyc_groups.get(acct, raw.iloc[0:0])
        pg = pay_groups.get(acct, pay.iloc[0:0])
        vg = vps_groups.get(acct, vps_rows.iloc[0:0])

        kyc_names = unique_join(kg.get("PAYMETER - Customer Name", pd.Series(dtype=object)))
        kyc_meters = unique_join(kg.get("METER NUMBER", pd.Series(dtype=object)))
        kyc_mac = unique_join(kg.get("MACRON ACCOUNT NUMBER", pd.Series(dtype=object)))
        kyc_mac_list = split_multi(kyc_mac)

        observed_disco = unique_join(pg.get("Disco Account Number", pd.Series(dtype=object)))
        observed_meters = unique_join(pg.get("Meter Number", pd.Series(dtype=object)))
        observed_names = unique_join(pg.get("Customer Name", pd.Series(dtype=object)))
        phones = unique_join(pg.get("Phone Number", pd.Series(dtype=object)))

        mac_meter_values = []
        mac_tx_count = 0
        for ma in kyc_mac_list:
            mg = mac_groups.get(ma)
            if mg is not None:
                mac_tx_count += len(mg)
                mac_meter_values.extend(mg.get("METER NUMBER", pd.Series(dtype=object)).tolist())
        mac_meters = unique_join(mac_meter_values)

        conflicts = []
        if len(kyc_mac_list) > 1:
            conflicts.append("Multiple Macron accounts in uploaded KYC for one Paymeter account")
        for ma in kyc_mac_list:
            if len(uploaded_mac_to_pay.get(ma, set())) > 1:
                conflicts.append("Macron account is assigned to multiple Paymeter accounts in uploaded KYC")
                break
        observed_disco_list = split_multi(observed_disco)
        if len(observed_disco_list) > 1:
            conflicts.append("Multiple Paymeter Disco Account Numbers observed")
        if len(kyc_mac_list) == 1 and observed_disco_list and any(x != kyc_mac_list[0] for x in observed_disco_list):
            conflicts.append("Uploaded KYC Macron account disagrees with Paymeter Disco Account Number")

        in_uploaded_kyc = len(kg) > 0
        if conflicts:
            status = "KYC CONFLICT"
        elif in_uploaded_kyc and len(kyc_mac_list) == 1:
            status = "MASTER MAPPED"
        elif in_uploaded_kyc:
            status = "INCOMPLETE KYC"
        else:
            status = "NOT IN UPLOADED KYC"

        customer_name = most_common_text(kg.get("PAYMETER - Customer Name", pd.Series(dtype=object)))
        if not customer_name:
            customer_name = most_common_text(pg.get("Customer Name", pd.Series(dtype=object)))
        if not customer_name:
            customer_name = most_common_text(vg.get("VPS_source_acct_name", pd.Series(dtype=object)))

        records.append({
            "KYC_ID": acct,
            "Customer Name": customer_name,
            "VPS Virtual Account Number": acct,
            "VPS Source Account Name(s)": unique_join(vg.get("VPS_source_acct_name", pd.Series(dtype=object))),
            "VPS Source Account Number(s)": unique_join(vg.get("VPS_source_acct_no", pd.Series(dtype=object))),
            "Paymeter Account Number(s)": acct,
            "Paymeter Customer Name(s)": observed_names or kyc_names,
            "Customer Phone Number(s)": phones,
            "KYC Meter Number(s)": kyc_meters,
            "Paymeter Meter Number(s)": observed_meters,
            "Paymeter Disco Account Number(s)": observed_disco,
            "Macron Account Number(s)": kyc_mac,
            "Macron Meter Number(s)": mac_meters,
            "Uploaded KYC Row Count": len(kg),
            "VPS Transaction Count": len(vg),
            "Paymeter Transaction Count": len(pg),
            "Macron Transaction Count": mac_tx_count,
            "KYC Link Method": "Uploaded KYC master: Paymeter Account Number -> Macron Account Number" if in_uploaded_kyc else "Observed transaction data only",
            "KYC Conflict": "; ".join(conflicts),
            "KYC Status": status,
            "KYC Match Usable?": "YES" if status == "MASTER MAPPED" else "NO",
        })
    return pd.DataFrame(records)


def _unique_map_from_multivalue(kyc, column, normalizer=clean_id, usable_only=False):
    holder = defaultdict(set)
    for _, r in kyc.iterrows():
        if usable_only and clean_id(r.get("KYC Match Usable?")) != "YES":
            continue
        kid = clean_id(r.get("KYC_ID"))
        for item in split_multi(r.get(column)):
            key = normalizer(item)
            if key and kid:
                holder[key].add(kid)
    return {k: next(iter(v)) for k, v in holder.items() if len(v) == 1}


def build_kyc_maps(kyc):
    kyc_by_id = {clean_id(r["KYC_ID"]): r for _, r in kyc.iterrows() if clean_id(r["KYC_ID"])}
    pay_account = _unique_map_from_multivalue(kyc, "Paymeter Account Number(s)")
    mac_account = _unique_map_from_multivalue(kyc, "Macron Account Number(s)", usable_only=True)
    return kyc_by_id, pay_account, mac_account


def _apply_kyc_fields(rec, kid, kyc_by_id, method):
    kr = kyc_by_id.get(clean_id(kid))
    rec["KYC_ID"] = clean_id(kid) if kr is not None else ""
    rec["KYC_Match_Method"] = method if kr is not None else "NOT LINKED"
    rec["KYC_Customer_Name"] = clean_id(kr.get("Customer Name")) if kr is not None else ""
    rec["KYC_VPS_Virtual_Account"] = clean_id(kr.get("VPS Virtual Account Number")) if kr is not None else ""
    rec["KYC_Paymeter_Account(s)"] = clean_id(kr.get("Paymeter Account Number(s)")) if kr is not None else ""
    rec["KYC_Macron_Account(s)"] = clean_id(kr.get("Macron Account Number(s)")) if kr is not None else ""
    rec["KYC_Meter_Number(s)"] = clean_id(kr.get("Macron Meter Number(s)")) if kr is not None else ""
    rec["KYC_Status"] = clean_id(kr.get("KYC Status")) if kr is not None else "NO KYC"
    rec["KYC_Conflict"] = clean_id(kr.get("KYC Conflict")) if kr is not None else ""
    return rec


# ================================================================
# 4. ATTACH MASTER KYC TO PAYMETER-MACRON HISTORY
# ================================================================

def attach_kyc_to_pm_mac(pm_mac_merge, kyc):
    kyc_by_id, pay_map, mac_map = build_kyc_maps(kyc)
    rows = []
    for r in pm_mac_merge.to_dict("records"):
        pa = clean_id(r.get("PM_Account Number"))
        ma = clean_id(r.get("MAC_ACCOUNT NUMBER"))
        kid, method = "", ""
        if pa and pa in pay_map:
            kid, method = pay_map[pa], "Paymeter Account Number -> uploaded KYC"
        elif ma and ma in mac_map:
            kid, method = mac_map[ma], "Macron Account Number -> uploaded KYC"
        rows.append(_apply_kyc_fields(dict(r), kid, kyc_by_id, method))
    return pd.DataFrame(rows)


# ================================================================
# 5. ATTACH MASTER KYC TO VPS-BANK HISTORY
# ================================================================

def attach_kyc_to_vps_bank(vps_bank_merge, kyc):
    kyc_by_id, pay_map, _ = build_kyc_maps(kyc)
    rows = []
    for r in vps_bank_merge.to_dict("records"):
        va = clean_id(r.get("VPS_virtual_acct_no"))
        kid = pay_map.get(va, "") if va else ""
        rows.append(_apply_kyc_fields(dict(r), kid, kyc_by_id, "VPS virtual_acct_no -> uploaded KYC" if kid else ""))
    return pd.DataFrame(rows)


# ================================================================
# 6. VPS -> PAYMETER TRANSACTION BRIDGE
# ================================================================

def _near_tie(sorted_candidates, tolerance_minutes=0.25):
    if len(sorted_candidates) < 2:
        return False
    return abs(float(sorted_candidates[1][0]) - float(sorted_candidates[0][0])) <= tolerance_minutes


def build_vps_paymeter_bridge(vps_bank_kyc, pay, max_delay_hours=TRANSACTION_TIME_TOLERANCE_HOURS, high_confidence_minutes=HIGH_CONFIDENCE_MINUTES):
    pay_records = pay.to_dict("index")
    index = defaultdict(list)
    for pi, p in pay_records.items():
        pa = clean_id(p.get("Account Number"))
        gross = p.get("Input Amount")
        if pa and pd.notna(gross):
            index[(pa, round(float(gross), 2))].append(pi)

    used_pay = set()
    rows = []
    max_minutes = float(max_delay_hours) * 60.0
    vps_rows = vps_bank_kyc[vps_bank_kyc["VPS_VPS_Row"].notna()].copy()
    for _, l in vps_rows.iterrows():
        va = clean_id(l.get("VPS_virtual_acct_no"))
        gross = l.get("VPS_transaction_amount_minor")
        vdt = l.get("VPS_VPS_DateTime")
        chosen = None
        candidates = []
        if va and pd.notna(gross) and pd.notna(vdt):
            for pi in index.get((va, round(float(gross), 2)), []):
                if pi in used_pay:
                    continue
                pdt = pay_records[pi].get("PM_DateTime")
                if pd.isna(pdt):
                    continue
                gap = abs((pd.Timestamp(pdt) - pd.Timestamp(vdt)).total_seconds()) / 60.0
                if gap <= max_minutes:
                    candidates.append((gap, pi))
            candidates.sort()
            if candidates and not _near_tie(candidates):
                chosen = candidates[0]

        rec = {
            "Bridge_Row": len(rows),
            "Bridge_Match_Status": "MATCHED" if chosen else ("AMBIGUOUS - REVIEW" if candidates else "VPS ONLY - NO PAYMETER"),
            "Bridge_Match_Method": "VPS virtual account + Paymeter Account Number + exact gross + nearest time" if chosen else "",
            "Bridge_Confidence": ("HIGH" if chosen and chosen[0] <= high_confidence_minutes else "REVIEW") if chosen else "UNMATCHED",
            "Bridge_Time_Gap_Minutes": chosen[0] if chosen else np.nan,
            "Bridge_Review_Notes": "Multiple near-equal Paymeter candidates" if candidates and chosen is None else "",
        }
        rec.update(l.to_dict())
        if chosen:
            pi = chosen[1]
            used_pay.add(pi)
            rec.update(prefixed_record(pay_records[pi], "PM_"))
        rows.append(rec)

    for pi, p in pay_records.items():
        if pi in used_pay:
            continue
        rec = {
            "Bridge_Row": len(rows),
            "Bridge_Match_Status": "PAYMETER ONLY - NO VPS",
            "Bridge_Match_Method": "",
            "Bridge_Confidence": "UNMATCHED",
            "Bridge_Time_Gap_Minutes": np.nan,
            "Bridge_Review_Notes": "",
        }
        rec.update(prefixed_record(p, "PM_"))
        rows.append(rec)

    out = pd.DataFrame(rows)
    for c in ["VPS_VPS_Row", "VPS_virtual_acct_no", "PM_PM_Row", "PM_Account Number", "PM_Disco Account Number", "PM_Reference", "PM_Input Amount", "PM_Transaction Amount", "PM_PM_DateTime"]:
        if c not in out.columns:
            out[c] = np.nan
    return out


def _right_record(pay_record, mac_record, kyc_by_id, pay_map, mac_map):
    rec = {}
    if pay_record is not None:
        rec.update(prefixed_record(pay_record, "PM_"))
    if mac_record is not None:
        rec.update(prefixed_record(mac_record, "MAC_"))
    kid, method = "", ""
    pa = clean_id(pay_record.get("Account Number")) if pay_record is not None else ""
    ma = clean_id(mac_record.get("ACCOUNT NUMBER")) if mac_record is not None else ""
    if pa and pa in pay_map:
        kid, method = pay_map[pa], "Paymeter Account Number -> uploaded KYC"
    elif ma and ma in mac_map:
        kid, method = mac_map[ma], "Macron Account Number -> uploaded KYC"
    return _apply_kyc_fields(rec, kid, kyc_by_id, method)


def _usable_kyc_mac(l):
    if clean_id(l.get("KYC_Status")) != "MASTER MAPPED":
        return ""
    vals = split_multi(l.get("KYC_Macron_Account(s)"))
    return vals[0] if len(vals) == 1 else ""


def _pick_time_candidate(candidates, high_confidence_minutes=HIGH_CONFIDENCE_MINUTES):
    candidates = sorted(candidates)
    if not candidates or _near_tie(candidates):
        return None, "AMBIGUOUS" if candidates else "NONE"
    gap, idx = candidates[0]
    return (gap, idx), ("HIGH" if gap <= high_confidence_minutes else "REVIEW")


# ================================================================
# 7. MAIN RECONCILIATION: VPS AGAINST MACRON
# ================================================================

def reconcile_merged_reports(vps_bank_kyc, vps_pay_bridge, pay, mac, kyc,
                             max_delay_hours=TRANSACTION_TIME_TOLERANCE_HOURS,
                             high_confidence_minutes=HIGH_CONFIDENCE_MINUTES):
    """Reconcile VPS to Macron in strict evidence tiers.

    All exact Paymeter->Macron reference matches are allocated FIRST across the
    whole dataset. Only remaining VPS/Macron rows are eligible for account/value/time
    or account/time fallback. This prevents a time fallback from consuming a Macron
    row that belongs to another VPS transaction with exact reference evidence.
    """
    left_records = vps_bank_kyc.to_dict("index")
    pay_records = pay.to_dict("index")
    mac_records = mac.to_dict("index")
    kyc_by_id, pay_map, mac_map = build_kyc_maps(kyc)

    bridge_by_vps = {}
    for _, b in vps_pay_bridge.iterrows():
        vrow = b.get("VPS_VPS_Row")
        if pd.notna(vrow):
            bridge_by_vps[int(vrow)] = b.to_dict()

    mac_by_ref = defaultdict(list)
    mac_by_account = defaultdict(list)
    for mi, m in mac_records.items():
        ref = clean_id(m.get("REFERENCE ID")).casefold()
        acct = clean_id(m.get("ACCOUNT NUMBER"))
        if ref:
            mac_by_ref[ref].append(mi)
        if acct:
            mac_by_account[acct].append(mi)

    pay_by_ref = defaultdict(list)
    for pi, p in pay_records.items():
        ref = clean_id(p.get("Reference")).casefold()
        if ref:
            pay_by_ref[ref].append(pi)

    max_minutes = float(max_delay_hours) * 60.0
    vps_items = {}
    used_pay = set()
    for li, l in left_records.items():
        if pd.isna(l.get("VPS_VPS_Row")):
            continue
        vrow = int(l["VPS_VPS_Row"])
        bridge = bridge_by_vps.get(vrow, {})
        pidx = int(bridge["PM_PM_Row"]) if pd.notna(bridge.get("PM_PM_Row")) else None
        p = pay_records.get(pidx) if pidx is not None else None
        if pidx is not None:
            used_pay.add(pidx)
        identity_mac = _usable_kyc_mac(l)
        pm_disco = clean_id(p.get("Disco Account Number")) if p is not None else ""
        identity_note = ""
        if not identity_mac and pm_disco:
            identity_mac = pm_disco
            identity_note = "Macron account derived from Paymeter because uploaded KYC was not usable"
        vps_items[li] = {
            "l": l,
            "pidx": pidx,
            "p": p,
            "identity_mac": identity_mac,
            "pm_disco": pm_disco,
            "identity_note": identity_note,
        }

    vps_to_mac = {}
    match_meta = {}
    used_mac = set()

    # Tier 1: allocate all exact Paymeter Reference = Macron REFERENCE ID matches first.
    for li, item in vps_items.items():
        p = item["p"]
        if p is None:
            continue
        pref = clean_id(p.get("Reference")).casefold()
        if not pref:
            continue
        candidates = [mi for mi in mac_by_ref.get(pref, []) if mi not in used_mac]
        if len(candidates) != 1:
            if len(candidates) > 1:
                match_meta[li] = {
                    "method": "", "confidence": "UNMATCHED",
                    "notes": "Duplicate Macron reference candidates; reference not auto-assigned",
                    "time_gap": np.nan, "reference": "NO",
                }
            continue
        mi = candidates[0]
        m = mac_records[mi]
        identity_mac = item["identity_mac"]
        pm_disco = item["pm_disco"]
        acct_ok = not identity_mac or clean_id(m.get("ACCOUNT NUMBER")) == identity_mac
        if pm_disco:
            acct_ok = acct_ok and clean_id(m.get("ACCOUNT NUMBER")) == pm_disco
        gap = np.nan
        if pd.notna(item["l"].get("VPS_VPS_DateTime")) and pd.notna(m.get("MAC_DateTime")):
            gap = abs((pd.Timestamp(m["MAC_DateTime"]) - pd.Timestamp(item["l"]["VPS_VPS_DateTime"])).total_seconds()) / 60.0
        notes = item["identity_note"]
        if not acct_ok:
            notes = (notes + "; " if notes else "") + "Reference matched but account identity conflicts"
        vps_to_mac[li] = mi
        used_mac.add(mi)
        match_meta[li] = {
            "method": "VPS -> Paymeter account/gross/time; Paymeter Reference = Macron REFERENCE ID",
            "confidence": "HIGH" if acct_ok else "REVIEW",
            "notes": notes,
            "time_gap": gap,
            "reference": "YES",
        }

    # Tier 2: uploaded KYC Macron account + expected token value + time.
    edges = []
    candidates_by_vps = defaultdict(list)
    candidates_by_mac = defaultdict(list)
    for li, item in vps_items.items():
        if li in vps_to_mac:
            continue
        l = item["l"]
        identity_mac = item["identity_mac"]
        gross = l.get("VPS_transaction_amount_minor")
        vdt = l.get("VPS_VPS_DateTime")
        if not identity_mac or pd.isna(gross) or pd.isna(vdt):
            continue
        expected = float(gross) - MACRON_TOKEN_CHARGE
        for mi in mac_by_account.get(identity_mac, []):
            if mi in used_mac:
                continue
            m = mac_records[mi]
            if pd.isna(m.get("AMOUNT")) or not amount_equal(m.get("AMOUNT"), expected):
                continue
            mdt = m.get("MAC_DateTime")
            if pd.isna(mdt):
                continue
            gap = abs((pd.Timestamp(mdt) - pd.Timestamp(vdt)).total_seconds()) / 60.0
            if gap <= max_minutes:
                edges.append((gap, li, mi))
                candidates_by_vps[li].append((gap, mi))
                candidates_by_mac[mi].append((gap, li))

    for gap, li, mi in sorted(edges):
        if li in vps_to_mac or mi in used_mac:
            continue
        item = vps_items[li]
        unique = len(candidates_by_vps[li]) == 1 and len(candidates_by_mac[mi]) == 1
        confidence = "HIGH" if (unique and gap <= high_confidence_minutes and clean_id(item["l"].get("KYC_Status")) == "MASTER MAPPED") else "REVIEW"
        notes = item["identity_note"]
        if not unique:
            notes = (notes + "; " if notes else "") + "Multiple plausible account/value/time pairings; closest one-to-one pair selected"
        vps_to_mac[li] = mi
        used_mac.add(mi)
        match_meta[li] = {
            "method": "Uploaded KYC Macron Account + expected token value + time",
            "confidence": confidence,
            "notes": notes,
            "time_gap": gap,
            "reference": "NO",
        }

    # Tier 3: same account + nearest time only. Always REVIEW.
    time_edges = []
    time_by_vps = defaultdict(list)
    time_by_mac = defaultdict(list)
    for li, item in vps_items.items():
        if li in vps_to_mac:
            continue
        l = item["l"]
        identity_mac = item["identity_mac"]
        vdt = l.get("VPS_VPS_DateTime")
        if not identity_mac or pd.isna(vdt):
            continue
        vdt = pd.Timestamp(vdt)
        for mi in mac_by_account.get(identity_mac, []):
            if mi in used_mac:
                continue
            mdt = mac_records[mi].get("MAC_DateTime")
            if pd.isna(mdt):
                continue
            mdt = pd.Timestamp(mdt)
            if abs((mdt.date() - vdt.date()).days) > 1:
                continue
            gap = abs((mdt - vdt).total_seconds()) / 60.0
            if gap <= max_minutes:
                time_edges.append((gap, li, mi))
                time_by_vps[li].append((gap, mi))
                time_by_mac[mi].append((gap, li))

    for gap, li, mi in sorted(time_edges):
        if li in vps_to_mac or mi in used_mac:
            continue
        # Do not force an effectively tied time match.
        vchoices = sorted(time_by_vps[li])
        mchoices = sorted(time_by_mac[mi])
        if _near_tie(vchoices) or _near_tie(mchoices):
            continue
        item = vps_items[li]
        notes = item["identity_note"]
        notes = (notes + "; " if notes else "") + "Reference/expected-value match unavailable; time fallback used"
        if len(vchoices) > 1 or len(mchoices) > 1:
            notes += "; multiple same-account candidates existed, closest one-to-one pair selected"
        vps_to_mac[li] = mi
        used_mac.add(mi)
        match_meta[li] = {
            "method": "Uploaded KYC Macron Account + nearest time fallback",
            "confidence": "REVIEW",
            "notes": notes,
            "time_gap": gap,
            "reference": "NO",
        }

    rows = []
    # Output every VPS transaction after all matching tiers have been allocated.
    for li, item in vps_items.items():
        l = item["l"]
        p = item["p"]
        mi = vps_to_mac.get(li)
        m = mac_records.get(mi) if mi is not None else None
        meta = match_meta.get(li, {
            "method": "", "confidence": "UNMATCHED",
            "notes": item["identity_note"], "time_gap": np.nan, "reference": "NO",
        })
        r = _right_record(p, m, kyc_by_id, pay_map, mac_map) if (p is not None or m is not None) else None
        rows.append(_final_recon_row(l, r, meta))

    # Bank-only rows must remain visible.
    for li, l in left_records.items():
        if pd.notna(l.get("VPS_VPS_Row")):
            continue
        rows.append(_final_recon_row(l, None, {
            "method": "", "confidence": "UNMATCHED",
            "notes": "Bank movement not paired to VPS", "time_gap": np.nan, "reference": "NO",
        }))

    # Unmatched Macron rows; attach an unused Paymeter row by exact reference where available.
    for mi, m in mac_records.items():
        if mi in used_mac:
            continue
        p = None
        ref = clean_id(m.get("REFERENCE ID")).casefold()
        candidates = [x for x in pay_by_ref.get(ref, []) if x not in used_pay] if ref else []
        if len(candidates) == 1:
            pidx = candidates[0]
            p = pay_records[pidx]
            used_pay.add(pidx)
        r = _right_record(p, m, kyc_by_id, pay_map, mac_map)
        rows.append(_final_recon_row(None, r, {
            "method": "Paymeter Reference = Macron REFERENCE ID" if p is not None else "",
            "confidence": "REVIEW" if p is not None else "UNMATCHED",
            "notes": "Macron transaction has no paired VPS payment",
            "time_gap": np.nan,
            "reference": "YES" if p is not None else "NO",
        }))

    # Remaining Paymeter rows stay visible and count in source tie-out.
    for pi, p in pay_records.items():
        if pi in used_pay:
            continue
        r = _right_record(p, None, kyc_by_id, pay_map, mac_map)
        rows.append(_final_recon_row(None, r, {
            "method": "", "confidence": "UNMATCHED",
            "notes": "Paymeter transaction not paired to VPS/Macron",
            "time_gap": np.nan, "reference": "NO",
        }))

    out = pd.DataFrame(rows)
    out.insert(0, "FINAL_Row_ID", [f"REC-{i+1:07d}" for i in range(len(out))])
    return out


def _safe_diff(a, b):
    return float(a) - float(b) if pd.notna(a) and pd.notna(b) else np.nan


def _final_recon_row(l, r, match_meta):
    match_meta = match_meta or {}
    has_vps = l is not None and clean_id(l.get("VPS_VPS_Row")) != ""
    has_bank = l is not None and clean_id(l.get("BANK_BANK_Row")) != ""
    has_pm = r is not None and clean_id(r.get("PM_PM_Row")) != ""
    has_mac = r is not None and clean_id(r.get("MAC_MAC_Row")) != ""

    left_kyc = clean_id(l.get("KYC_ID")) if l else ""
    right_kyc = clean_id(r.get("KYC_ID")) if r else ""
    kyc = left_kyc or right_kyc
    left_kyc_status = clean_id(l.get("KYC_Status")) if l else ""
    right_kyc_status = clean_id(r.get("KYC_Status")) if r else ""
    if "CONFLICT" in left_kyc_status.upper() or "CONFLICT" in right_kyc_status.upper():
        kyc_status = "KYC CONFLICT"
    elif left_kyc and right_kyc:
        kyc_status = "KYC ALIGNED" if left_kyc == right_kyc else "KYC MISMATCH"
    elif kyc:
        kyc_status = "KYC ON ONE SIDE"
    else:
        kyc_status = "NO KYC"

    vps_gross = l.get("VPS_transaction_amount_minor") if has_vps else np.nan
    vps_settled = l.get("VPS_settled_amount_minor") if has_vps else np.nan
    bank_credit = l.get("BANK_Credit") if has_bank else np.nan
    bank_debit = l.get("BANK_Debit") if has_bank else np.nan
    pm_input = r.get("PM_Input Amount") if has_pm else np.nan
    pm_token = r.get("PM_Transaction Amount") if has_pm else np.nan
    mac_token = r.get("MAC_AMOUNT") if has_mac else np.nan
    expected_mac = float(vps_gross) - MACRON_TOKEN_CHARGE if pd.notna(vps_gross) else np.nan

    d_bank_vps = _safe_diff(bank_credit, vps_settled)
    d_vps_pm = _safe_diff(vps_gross, pm_input)
    d_pm_mac = _safe_diff(pm_token, mac_token)
    d_expected_mac = _safe_diff(expected_mac, mac_token)

    if has_vps and has_mac and pd.notna(d_expected_mac):
        amount_status = "AMOUNTS ALIGNED" if abs(d_expected_mac) <= AMOUNT_TOLERANCE else "AMOUNT DIFFERENCE"
    else:
        amount_status = "NOT COMPARABLE"

    support_issues = []
    if has_bank and has_vps and pd.notna(d_bank_vps) and abs(d_bank_vps) > AMOUNT_TOLERANCE:
        support_issues.append("Bank credit differs from VPS settled amount")
    if has_vps and has_pm and pd.notna(d_vps_pm) and abs(d_vps_pm) > AMOUNT_TOLERANCE:
        support_issues.append("VPS gross differs from Paymeter input")
    if has_pm and has_mac and pd.notna(d_pm_mac) and abs(d_pm_mac) > AMOUNT_TOLERANCE:
        support_issues.append("Paymeter token differs from Macron amount")
    support_status = "SUPPORTING CHECKS ALIGNED" if not support_issues else "SUPPORTING DIFFERENCE"

    if has_vps and not has_mac:
        status = "CUSTOMER PAID / VPS RECEIVED - NO MACRON TOKEN"
    elif has_mac and not has_vps:
        status = "MACRON TOKEN - NO VPS PAYMENT"
    elif has_vps and has_mac and kyc_status in {"KYC MISMATCH", "KYC CONFLICT"}:
        status = kyc_status
    elif has_vps and has_mac and amount_status == "AMOUNT DIFFERENCE":
        status = "VPS / MACRON VALUE DIFFERENCE"
    elif has_vps and has_mac and not has_bank:
        status = "VPS / MACRON MATCH - BANK MISSING"
    elif has_vps and has_mac and amount_status == "AMOUNTS ALIGNED":
        status = "FULLY RECONCILED"
    elif has_vps and not has_bank:
        status = "VPS PRESENT - BANK MISSING"
    elif has_pm and not has_vps and not has_mac:
        status = "PAYMETER ONLY - NO VPS / MACRON"
    elif has_bank and not has_vps:
        if pd.notna(bank_debit):
            status = "BANK DEBIT / OUTFLOW - NO VPS"
        elif pd.notna(bank_credit):
            status = "BANK CREDIT / INFLOW - NO VPS"
        else:
            status = "BANK MOVEMENT - NO VPS"
    else:
        status = "PARTIAL / REVIEW"

    rec = {
        "FINAL_Reconciliation_Status": status,
        "FINAL_Match_Method": clean_id(match_meta.get("method")),
        "FINAL_Match_Confidence": clean_id(match_meta.get("confidence")),
        "FINAL_Review_Notes": clean_id(match_meta.get("notes")),
        "FINAL_Reference_Confirmation": clean_id(match_meta.get("reference")),
        "FINAL_Time_Gap_Minutes": match_meta.get("time_gap", np.nan),
        "FINAL_KYC_ID": kyc,
        "FINAL_KYC_Status": kyc_status,
        "FINAL_Customer_Name": (clean_id(l.get("KYC_Customer_Name")) if l else "") or (clean_id(r.get("KYC_Customer_Name")) if r else "") or (clean_id(r.get("PM_Customer Name")) if r else ""),
        "Bank Present?": "YES" if has_bank else "NO",
        "VPS Present?": "YES" if has_vps else "NO",
        "Paymeter Present?": "YES" if has_pm else "NO",
        "Macron Present?": "YES" if has_mac else "NO",
        "VPS Customer Amount Paid": vps_gross,
        "Expected Macron Token (VPS Paid - 100)": expected_mac,
        "Bank Credit": bank_credit,
        "Bank Debit": bank_debit,
        "VPS Settled Amount": vps_settled,
        "Paymeter Input Amount": pm_input,
        "Paymeter Token/Net Amount": pm_token,
        "Macron Token Amount": mac_token,
        "Difference Bank Credit - VPS Settled": d_bank_vps,
        "Difference VPS Paid - Paymeter Input": d_vps_pm,
        "Difference Paymeter Token - Macron Token": d_pm_mac,
        "Difference Expected Macron - Actual Macron": d_expected_mac,
        "FINAL_Amount_Status": amount_status,
        "FINAL_Supporting_Check_Status": support_status,
        "FINAL_Supporting_Check_Notes": "; ".join(support_issues),
    }
    if l:
        rec.update(prefixed_record(l, "LEFT_"))
    if r:
        rec.update(prefixed_record(r, "RIGHT_"))
    return rec


# ================================================================
# 8. VPS AGAINST MACRON FOCUSED VIEW
# ================================================================

def build_vps_macron_comparison(final_recon):
    cols = [
        "FINAL_Row_ID", "FINAL_KYC_ID", "FINAL_Customer_Name", "FINAL_KYC_Status",
        "LEFT_VPS_virtual_acct_no", "LEFT_VPS_session_id", "LEFT_VPS_settlement_ref",
        "RIGHT_PM_Account Number", "RIGHT_PM_Disco Account Number", "RIGHT_PM_Reference",
        "RIGHT_MAC_ACCOUNT NUMBER", "RIGHT_MAC_METER NUMBER", "RIGHT_MAC_REFERENCE ID",
        "VPS Customer Amount Paid", "Expected Macron Token (VPS Paid - 100)",
        "Macron Token Amount", "Difference Expected Macron - Actual Macron",
        "FINAL_Reference_Confirmation", "FINAL_Time_Gap_Minutes", "FINAL_Match_Confidence",
        "VPS Present?", "Macron Present?", "FINAL_Match_Method", "FINAL_Reconciliation_Status",
    ]
    work = final_recon.copy()
    for c in cols:
        if c not in work.columns:
            work[c] = np.nan
    out = work[(work["VPS Present?"] == "YES") | (work["Macron Present?"] == "YES")][cols].copy()
    out["VPS_vs_Macron_Status"] = np.where(
        (out["VPS Present?"] == "YES") & (out["Macron Present?"] == "NO"),
        "PAID / NO TOKEN",
        np.where(
            (out["VPS Present?"] == "NO") & (out["Macron Present?"] == "YES"),
            "TOKEN / NO VPS PAYMENT",
            np.where(
                pd.to_numeric(out["Difference Expected Macron - Actual Macron"], errors="coerce").abs().fillna(np.inf) <= AMOUNT_TOLERANCE,
                "VALUE MATCH",
                "VALUE DIFFERENCE",
            ),
        ),
    )
    return out


def build_posting_ledger(final_recon):
    """Seven-column posting view requested by the user.

    Reconciled customer flow keeps the original accounting meanings:
      Cash In  = VPS gross less the fixed NGN 100 customer charge.
      Cash Out = actual Macron amount issued.
      Interest = NGN 100 less actual bank charge (VPS gross - VPS settled).

    Genuine Bank-only movements are also retained. For those rows, Cash In/Out
    is the actual bank Credit/Debit and customer/account fields remain blank.
    """
    rows = []
    for _, r in final_recon.iterrows():
        vps_present = clean_id(r.get("VPS Present?")) == "YES"
        mac_present = clean_id(r.get("Macron Present?")) == "YES"
        bank_present = clean_id(r.get("Bank Present?")) == "YES"
        pm_present = clean_id(r.get("Paymeter Present?")) == "YES"

        trx_date = pd.NaT
        # Posting date follows the actual transaction side: VPS first for
        # customer receipts, then Macron/Paymeter, and Bank for bank-only rows.
        for key in ["LEFT_VPS_VPS_DateTime", "RIGHT_MAC_MAC_DateTime", "RIGHT_PM_PM_DateTime", "LEFT_BANK_BANK_DateTime"]:
            value = r.get(key)
            if pd.notna(value):
                trx_date = pd.Timestamp(value).normalize()
                break

        bank_only = bank_present and not vps_present and not mac_present and not pm_present
        if bank_only:
            bank_credit = pd.to_numeric(pd.Series([r.get("LEFT_BANK_Credit")]), errors="coerce").iloc[0]
            bank_debit = pd.to_numeric(pd.Series([r.get("LEFT_BANK_Debit")]), errors="coerce").iloc[0]
            rows.append({
                "Transaction Date": trx_date,
                "Customer Name": "",
                "Account Number": "",
                "Cash In": float(bank_credit) if pd.notna(bank_credit) else 0.0,
                "Cash Out": float(bank_debit) if pd.notna(bank_debit) else 0.0,
                "Interest": 0.0,
                "Source": "Bank",
            })
            continue

        gross = r.get("VPS Customer Amount Paid")
        settled = r.get("VPS Settled Amount")
        cash_in = float(gross) - MACRON_TOKEN_CHARGE if vps_present and pd.notna(gross) else 0.0
        cash_out = float(r.get("Macron Token Amount")) if mac_present and pd.notna(r.get("Macron Token Amount")) else 0.0
        if vps_present and pd.notna(gross) and pd.notna(settled):
            actual_bank_charge = float(gross) - float(settled)
            interest = MACRON_TOKEN_CHARGE - actual_bank_charge
        else:
            interest = 0.0

        if bank_present and vps_present and mac_present:
            source = "Bank/VPS/Macron"
        elif bank_present and vps_present:
            source = "Bank/VPS"
        elif vps_present and mac_present:
            source = "VPS/Macron"
        elif mac_present:
            source = "Macron"
        elif vps_present:
            source = "VPS"
        elif pm_present:
            source = "Paymeter"
        elif bank_present:
            source = "Bank"
        else:
            source = ""

        rows.append({
            "Transaction Date": trx_date,
            "Customer Name": clean_id(r.get("FINAL_Customer_Name")),
            "Account Number": clean_id(r.get("FINAL_KYC_ID")) or clean_id(r.get("LEFT_VPS_virtual_acct_no")) or clean_id(r.get("RIGHT_PM_Account Number")),
            "Cash In": cash_in,
            "Cash Out": cash_out,
            "Interest": interest,
            "Source": source,
        })
    return pd.DataFrame(rows)


# ================================================================
# 8. COMPLETE CUSTOMER SUMMARY — NO SOURCE FIGURE MAY DISAPPEAR
# ================================================================

def _summary_key(r):
    kid = clean_id(r.get("FINAL_KYC_ID"))
    if kid:
        return f"KYC::{kid}", "KYC CUSTOMER"
    pa = clean_id(r.get("RIGHT_PM_Account Number"))
    if pa:
        return f"UNLINKED-PAYMETER::{pa}", "UNLINKED PAYMETER"
    ma = clean_id(r.get("RIGHT_MAC_ACCOUNT NUMBER"))
    if ma:
        return f"UNLINKED-MACRON::{ma}", "UNLINKED MACRON"
    mm = normalize_meter(r.get("RIGHT_MAC_METER NUMBER"))
    if mm:
        return f"UNLINKED-MACRON-METER::{mm}", "UNLINKED MACRON METER"
    if clean_id(r.get("Bank Present?")) == "YES":
        return "UNIDENTIFIED-BANK-CREDIT", "UNIDENTIFIED BANK CREDIT"
    return f"UNIDENTIFIED::{clean_id(r.get('FINAL_Row_ID'))}", "UNIDENTIFIED"


def build_customer_summary(final_recon, kyc):
    work = final_recon.copy()
    # Bank-only debits/outflows belong in the Cash Posting history, not in
    # customer exposure or the VPS-receipt Customer Summary.
    bank_debit = pd.to_numeric(work.get("Bank Debit", pd.Series(index=work.index, dtype=float)), errors="coerce").notna()
    debit_only = (work["Bank Present?"] == "YES") & (work["VPS Present?"] == "NO") & (work["Paymeter Present?"] == "NO") & (work["Macron Present?"] == "NO") & bank_debit
    work = work[~debit_only].copy()
    keys = work.apply(_summary_key, axis=1)
    work["Summary Customer Key"] = [x[0] for x in keys]
    work["Customer Classification"] = [x[1] for x in keys]
    kyc_lookup = {clean_id(r["KYC_ID"]): r for _, r in kyc.iterrows() if clean_id(r["KYC_ID"])}
    rows = []

    for key, g in work.groupby("Summary Customer Key", sort=True):
        kid = most_common_text(g["FINAL_KYC_ID"])
        kr = kyc_lookup.get(kid)
        vps_count = int((g["VPS Present?"] == "YES").sum())
        row = {
            "Summary Customer Key": key,
            "Customer Classification": most_common_text(g["Customer Classification"]),
            "KYC_ID": kid,
            "Customer Name": clean_id(kr.get("Customer Name")) if kr is not None else (most_common_text(g["FINAL_Customer_Name"]) or "UNIDENTIFIED CUSTOMER"),
            "VPS Virtual Account Number": clean_id(kr.get("VPS Virtual Account Number")) if kr is not None else unique_join(g.get("LEFT_VPS_virtual_acct_no", pd.Series(dtype=object))),
            "Paymeter Account Number(s)": clean_id(kr.get("Paymeter Account Number(s)")) if kr is not None else unique_join(g.get("RIGHT_PM_Account Number", pd.Series(dtype=object))),
            "Paymeter Meter Number(s)": clean_id(kr.get("Paymeter Meter Number(s)")) if kr is not None else unique_join(g.get("RIGHT_PM_Meter Number", pd.Series(dtype=object))),
            "Macron Account Number(s)": clean_id(kr.get("Macron Account Number(s)")) if kr is not None else unique_join(g.get("RIGHT_MAC_ACCOUNT NUMBER", pd.Series(dtype=object))),
            "Macron Meter Number(s)": clean_id(kr.get("Macron Meter Number(s)")) if kr is not None else unique_join(g.get("RIGHT_MAC_METER NUMBER", pd.Series(dtype=object))),
            "Total Bank Credit Count": int(pd.to_numeric(g["Bank Credit"], errors="coerce").notna().sum()),
            "Total VPS Payment Count": vps_count,
            "Total Paymeter Transaction Count": int((g["Paymeter Present?"] == "YES").sum()),
            "Total Macron Vend Count": int((g["Macron Present?"] == "YES").sum()),
            "Total Bank Credit": pd.to_numeric(g["Bank Credit"], errors="coerce").sum(),
            "Total VPS Customer Amount Paid": pd.to_numeric(g["VPS Customer Amount Paid"], errors="coerce").sum(),
            "Total VPS Settled Amount": pd.to_numeric(g["VPS Settled Amount"], errors="coerce").sum(),
            "Expected Total Macron Token": pd.to_numeric(g["Expected Macron Token (VPS Paid - 100)"], errors="coerce").sum(),
            "Total Paymeter Input Amount": pd.to_numeric(g["Paymeter Input Amount"], errors="coerce").sum(),
            "Total Paymeter Token/Net Amount": pd.to_numeric(g["Paymeter Token/Net Amount"], errors="coerce").sum(),
            "Total Macron Token Amount": pd.to_numeric(g["Macron Token Amount"], errors="coerce").sum(),
            "VPS Paid / No Token Count": int(((g["VPS Present?"] == "YES") & (g["Macron Present?"] == "NO")).sum()),
            "Token / No VPS Payment Count": int(((g["Macron Present?"] == "YES") & (g["VPS Present?"] == "NO")).sum()),
            "Difference Expected Macron - Actual Macron": pd.to_numeric(g["Expected Macron Token (VPS Paid - 100)"], errors="coerce").sum() - pd.to_numeric(g["Macron Token Amount"], errors="coerce").sum(),
            "Difference VPS Paid - Paymeter Input": pd.to_numeric(g["VPS Customer Amount Paid"], errors="coerce").sum() - pd.to_numeric(g["Paymeter Input Amount"], errors="coerce").sum(),
            "Difference Paymeter Token - Macron Token": pd.to_numeric(g["Paymeter Token/Net Amount"], errors="coerce").sum() - pd.to_numeric(g["Macron Token Amount"], errors="coerce").sum(),
            "KYC Status": clean_id(kr.get("KYC Status")) if kr is not None else "UNLINKED / NO FORMAL KYC",
        }
        rows.append(row)
    return pd.DataFrame(rows)


def build_summary_control(summary, bank, vps, pay, mac):
    def s(series):
        return pd.to_numeric(series, errors="coerce").sum()
    controls = [
        ("Bank", "Credit Count", int(bank["Credit"].notna().sum()), int(pd.to_numeric(summary["Total Bank Credit Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("Bank", "Credit Amount", s(bank.loc[bank["Credit"].notna(), "Credit"]), s(summary["Total Bank Credit"]), "AMOUNT"),
        ("VPS", "Transaction Count", len(vps), int(pd.to_numeric(summary["Total VPS Payment Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("VPS", "transaction_amount_minor", s(vps["transaction_amount_minor"]), s(summary["Total VPS Customer Amount Paid"]), "AMOUNT"),
        ("VPS", "settled_amount_minor", s(vps["settled_amount_minor"]), s(summary["Total VPS Settled Amount"]), "AMOUNT"),
        ("Paymeter", "Transaction Count", len(pay), int(pd.to_numeric(summary["Total Paymeter Transaction Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("Paymeter", "Input Amount", s(pay["Input Amount"]), s(summary["Total Paymeter Input Amount"]), "AMOUNT"),
        ("Paymeter", "Transaction Amount", s(pay["Transaction Amount"]), s(summary["Total Paymeter Token/Net Amount"]), "AMOUNT"),
        ("Macron", "Vend Count", len(mac), int(pd.to_numeric(summary["Total Macron Vend Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("Macron", "AMOUNT", s(mac["AMOUNT"]), s(summary["Total Macron Token Amount"]), "AMOUNT"),
        ("VPS Derived", "Expected Macron Token = transaction_amount_minor - 100", s(vps["transaction_amount_minor"]) - MACRON_TOKEN_CHARGE * int(vps["transaction_amount_minor"].notna().sum()), s(summary["Expected Total Macron Token"]), "AMOUNT"),
    ]
    rows = []
    for source, metric, source_total, summary_total, typ in controls:
        diff = float(summary_total) - float(source_total)
        okay = abs(diff) < 0.5 if typ == "COUNT" else abs(diff) <= AMOUNT_TOLERANCE
        rows.append({"Source": source, "Metric": metric, "Source Total": source_total, "Customer Summary Total": summary_total, "Difference": diff, "Type": typ, "Status": "MATCH" if okay else "MISMATCH"})
    return pd.DataFrame(rows)


def validate_tieout(control):
    bad = control[control["Status"] != "MATCH"]
    if len(bad):
        detail = "; ".join(f"{r['Source']} {r['Metric']} diff={r['Difference']}" for _, r in bad.iterrows())
        raise ValueError("Customer Summary does not tie to source reports: " + detail)


# ================================================================
# MANAGEMENT / EXCEPTION REPORTS
# ================================================================


def _option2_frequency_rating(payment_count):
    """
    Customer purchase-frequency rating.

    This is intentionally separate from the financial exposure rating:
      Very Frequent = 13+ VPS purchases
      Frequent      = 7-12
      Regular       = 4-6
      Occasional    = 1-3
      No VPS History= 0
    """
    try:
        count = int(float(payment_count))
    except Exception:
        count = 0

    if count >= OPTION2_VERY_FREQUENT_MIN:
        return "Very Frequent"
    if count >= OPTION2_FREQUENT_MIN:
        return "Frequent"
    if count >= OPTION2_REGULAR_MIN:
        return "Regular"
    if count >= 1:
        return "Occasional"
    return "No VPS History"


def _option2_exposure_category(share):
    """
    Option 2 financial-materiality rating based on the individual
    customer's share of the RELEVANT total outstanding.

      Critical         >= 10%
      High Materiality  2% to <10%
      Significant       0.5% to <2%
      Low Materiality  < 0.5%
    """
    try:
        share = float(share)
    except Exception:
        share = 0.0

    if share >= OPTION2_CRITICAL_SHARE:
        return "Critical"
    if share >= OPTION2_HIGH_SHARE:
        return "High Materiality"
    if share >= OPTION2_SIGNIFICANT_SHARE:
        return "Significant"
    return "Low Materiality"


def build_option2_management(customer_summary):
    """
    Build the Option 2 management report directly from Customer Summary.

    Business logic:
      Expected Token = Total VPS Customer Amount Paid
                       - (Total VPS Payment Count * N100)

      If Actual Macron Token > Expected Token:
          Customer owes us the excess token value.

      If Expected Token > Actual Macron Token:
          We owe the customer token value.

    Financial category is based on the customer's INDIVIDUAL SHARE of
    the applicable total outstanding. Purchase frequency is a separate
    behavioural rating and does not change the exposure category.
    """
    required = [
        "Summary Customer Key",
        "Customer Classification",
        "Customer Name",
        "Total VPS Payment Count",
        "Total VPS Customer Amount Paid",
        "Total Macron Token Amount",
        "VPS Paid / No Token Count",
        "Token / No VPS Payment Count",
        "KYC Status",
    ]
    require_columns(customer_summary, required, "Customer Summary")

    work = customer_summary.copy()

    numeric_cols = [
        "Total VPS Payment Count",
        "Total VPS Customer Amount Paid",
        "Total Macron Token Amount",
        "VPS Paid / No Token Count",
        "Token / No VPS Payment Count",
    ]
    for c in numeric_cols:
        work[c] = pd.to_numeric(work[c], errors="coerce").fillna(0)

    # Recalculate expected token from the clarified business rule.
    work["Option2 Expected Token Value"] = (
        work["Total VPS Customer Amount Paid"]
        - (work["Total VPS Payment Count"] * MACRON_TOKEN_CHARGE)
    )
    work["Option2 Actual Macron Token"] = work["Total Macron Token Amount"]

    # Positive means too much token was issued; negative means token is still due.
    work["Option2 Actual Less Expected"] = (
        work["Option2 Actual Macron Token"]
        - work["Option2 Expected Token Value"]
    )

    work["Option2 Recovery Outstanding"] = work["Option2 Actual Less Expected"].clip(lower=0)
    work["Option2 Token Due"] = (-work["Option2 Actual Less Expected"]).clip(lower=0)

    work["Direction"] = np.select(
        [
            work["Option2 Recovery Outstanding"] > AMOUNT_TOLERANCE,
            work["Option2 Token Due"] > AMOUNT_TOLERANCE,
        ],
        [
            "Customer Owes Us",
            "We Owe Customer Token",
        ],
        default="Balanced",
    )

    # Only management exceptions belong in the Option 2 report.
    out = work[work["Direction"] != "Balanced"].copy()

    out["Outstanding Amount"] = np.where(
        out["Direction"].eq("Customer Owes Us"),
        out["Option2 Recovery Outstanding"],
        out["Option2 Token Due"],
    )

    direction_totals = (
        out.groupby("Direction")["Outstanding Amount"]
        .sum()
        .to_dict()
    )

    out["Relevant Total Outstanding"] = out["Direction"].map(direction_totals).fillna(0)
    out["Individual Share of Total"] = np.where(
        out["Relevant Total Outstanding"].abs() > AMOUNT_TOLERANCE,
        out["Outstanding Amount"] / out["Relevant Total Outstanding"],
        0.0,
    )

    out["Exposure Category"] = out["Individual Share of Total"].map(
        _option2_exposure_category
    )
    out["Purchase Frequency Rating"] = out["Total VPS Payment Count"].map(
        _option2_frequency_rating
    )

    # Rank separately on each side, highest outstanding first.
    out["Rank"] = (
        out.groupby("Direction")["Outstanding Amount"]
        .rank(method="first", ascending=False)
        .astype(int)
    )

    selected = [
        "Direction",
        "Rank",
        "Summary Customer Key",
        "Customer Name",
        "Customer Classification",
        "KYC Status",
        "Total VPS Payment Count",
        "Purchase Frequency Rating",
        "VPS Paid / No Token Count",
        "Token / No VPS Payment Count",
        "Option2 Expected Token Value",
        "Option2 Actual Macron Token",
        "Outstanding Amount",
        "Relevant Total Outstanding",
        "Individual Share of Total",
        "Exposure Category",
    ]

    out = out[selected].copy()
    out = out.sort_values(
        ["Direction", "Rank", "Customer Name"],
        ascending=[True, True, True],
    ).reset_index(drop=True)
    return out


def _option2_category_summary(option2_detail, direction):
    """Return category counts, amount and share for one Option 2 direction."""
    categories = ["Critical", "High Materiality", "Significant", "Low Materiality"]
    subset = option2_detail[option2_detail["Direction"] == direction].copy()
    total = pd.to_numeric(subset["Outstanding Amount"], errors="coerce").fillna(0).sum()

    rows = []
    for category in categories:
        part = subset[subset["Exposure Category"] == category]
        amount = pd.to_numeric(part["Outstanding Amount"], errors="coerce").fillna(0).sum()
        rows.append({
            "Exposure Category": category,
            "Customers": len(part),
            "Outstanding Amount": float(amount),
            "% of Direction Total": float(amount / total) if total else 0.0,
        })
    return pd.DataFrame(rows)


def _write_option2_management_sheet(writer, option2_detail):
    """
    Write a management-friendly Option 2 worksheet.

    The top of the tab is a simple management report.
    The full customer supporting detail is retained below it.
    """
    sheet_name = "16 Option 2 Management"
    wb = writer.book
    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws

    # -------------------------
    # Formats
    # -------------------------
    fmt_title = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": "#17365D",
        "font_size": 16, "align": "center", "valign": "vcenter",
    })
    fmt_subtitle = wb.add_format({
        "italic": True, "font_color": "#44546A", "bg_color": "#EAF2F8",
        "text_wrap": True, "valign": "vcenter",
    })
    fmt_section = wb.add_format({
        "bold": True, "font_color": "#17365D", "bg_color": "#D9EAF7",
        "border": 1,
    })
    fmt_header = wb.add_format({
        "bold": True, "font_color": "white", "bg_color": "#1F4E78",
        "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter",
    })
    fmt_money = wb.add_format({"num_format": '₦#,##0.00;[Red](₦#,##0.00);-', "border": 1})
    fmt_number = wb.add_format({"num_format": '#,##0', "border": 1})
    fmt_percent = wb.add_format({"num_format": '0.0%', "border": 1})
    fmt_text = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
    fmt_center = wb.add_format({"border": 1, "align": "center", "valign": "vcenter"})
    fmt_note = wb.add_format({
        "font_color": "#7F6000", "bg_color": "#FFF2CC",
        "text_wrap": True, "valign": "top", "border": 1,
    })
    fmt_kpi_label = wb.add_format({
        "bold": True, "font_color": "#17365D", "bg_color": "#D9EAF7",
        "align": "center", "valign": "vcenter", "border": 1,
    })
    fmt_kpi_money = wb.add_format({
        "bold": True, "font_size": 14, "font_color": "#17365D",
        "bg_color": "#EAF2F8", "align": "center", "valign": "vcenter",
        "num_format": '₦#,##0.00', "border": 1,
    })
    fmt_kpi_count = wb.add_format({
        "bold": True, "font_size": 14, "font_color": "#17365D",
        "bg_color": "#EAF2F8", "align": "center", "valign": "vcenter",
        "num_format": '#,##0', "border": 1,
    })
    fmt_critical = wb.add_format({"bg_color": "#F4CCCC", "font_color": "#9C0006", "bold": True, "border": 1})
    fmt_high = wb.add_format({"bg_color": "#FCE5CD", "font_color": "#C65911", "bold": True, "border": 1})
    fmt_significant = wb.add_format({"bg_color": "#FFF2CC", "font_color": "#7F6000", "bold": True, "border": 1})
    fmt_low = wb.add_format({"bg_color": "#E2F0D9", "font_color": "#375623", "border": 1})

    recovery = option2_detail[option2_detail["Direction"] == "Customer Owes Us"].copy()
    token_due = option2_detail[option2_detail["Direction"] == "We Owe Customer Token"].copy()

    recovery = recovery.sort_values(["Rank", "Customer Name"])
    token_due = token_due.sort_values(["Rank", "Customer Name"])

    recovery_total = pd.to_numeric(recovery["Outstanding Amount"], errors="coerce").fillna(0).sum()
    token_due_total = pd.to_numeric(token_due["Outstanding Amount"], errors="coerce").fillna(0).sum()
    net_position = recovery_total - token_due_total

    rec_summary = _option2_category_summary(option2_detail, "Customer Owes Us")
    due_summary = _option2_category_summary(option2_detail, "We Owe Customer Token")

    # -------------------------
    # Title
    # -------------------------
    ws.merge_range("A1:P1", "MANAGEMENT SUMMARY — OPTION 2 CUSTOMER TOKEN RECONCILIATION", fmt_title)
    ws.set_row(0, 28)
    ws.merge_range(
        "A2:P2",
        "Exposure category = each customer's individual share of the relevant total outstanding. "
        "Purchase frequency is shown separately and does not change financial materiality.",
        fmt_subtitle,
    )

    # -------------------------
    # KPI cards
    # -------------------------
    kpis = [
        ("A4:C4", "A5:C6", "CUSTOMERS OWING US", len(recovery), False),
        ("D4:F4", "D5:F6", "RECOVERY OUTSTANDING", recovery_total, True),
        ("G4:I4", "G5:I6", "CUSTOMERS WE OWE TOKEN", len(token_due), False),
        ("J4:L4", "J5:L6", "TOKEN VALUE DUE", token_due_total, True),
        ("M4:P4", "M5:P6", "NET RECEIVABLE POSITION", net_position, True),
    ]
    for label_range, value_range, label, value, is_money in kpis:
        ws.merge_range(label_range, label, fmt_kpi_label)
        ws.merge_range(value_range, value, fmt_kpi_money if is_money else fmt_kpi_count)

    # -------------------------
    # Rating guides
    # -------------------------
    ws.merge_range("A8:P8", "OPTION 2 EXPOSURE & PURCHASE-FREQUENCY GUIDE", fmt_section)

    exposure_guide = [
        ["Exposure Category", "Individual Share of Total", "Meaning", "Management Use"],
        ["Critical", ">=10%", "One customer individually represents at least 10% of total exposure", "Executive attention / immediate validation"],
        ["High Materiality", "2% to <10%", "Customer is individually material to the total exposure", "Prioritise for action"],
        ["Significant", "0.5% to <2%", "Meaningful individual exposure", "Monitor and resolve systematically"],
        ["Low Materiality", "<0.5%", "Small individual contribution to total exposure", "Routine resolution / batch handling"],
    ]
    for c, val in enumerate(exposure_guide[0]):
        ws.write(8, c, val, fmt_header)
    for r, row in enumerate(exposure_guide[1:], start=9):
        for c, val in enumerate(row):
            ws.write(r, c, val, fmt_text)

    frequency_guide = [
        ["Purchase Frequency", "VPS Purchases", "Interpretation", "Why It Matters"],
        ["Very Frequent", "13+", "Approx. top activity tier", "Active / high-touch customer"],
        ["Frequent", "7-12", "Upper activity range", "Regular commercial relationship"],
        ["Regular", "4-6", "Mid-range activity", "Established moderate activity"],
        ["Occasional", "1-3", "Lower purchase-frequency range", "Less frequent customer"],
    ]
    for c, val in enumerate(frequency_guide[0], start=7):
        ws.write(8, c, val, fmt_header)
    for r, row in enumerate(frequency_guide[1:], start=9):
        for c, val in enumerate(row, start=7):
            ws.write(r, c, val, fmt_text)

    ws.merge_range(
        "M9:P13",
        "How to read the report:\n\n"
        "Exposure Category = how much the customer matters financially.\n\n"
        "Purchase Frequency = how active the customer is.\n\n"
        "The two ratings are intentionally separate.",
        fmt_note,
    )

    # -------------------------
    # Distribution summaries
    # -------------------------
    ws.merge_range("A16:G16", "CUSTOMERS OWING US — EXPOSURE DISTRIBUTION", fmt_section)
    ws.merge_range("I16:P16", "CUSTOMERS WE OWE TOKEN — EXPOSURE DISTRIBUTION", fmt_section)

    rec_headers = ["Exposure Category", "Customers", "Outstanding (₦)", "% of Total"]
    due_headers = ["Exposure Category", "Customers", "Token Due (₦)", "% of Total"]
    for c, h in enumerate(rec_headers):
        ws.write(16, c, h, fmt_header)
    for c, h in enumerate(due_headers, start=8):
        ws.write(16, c, h, fmt_header)

    for i, row in rec_summary.iterrows():
        rr = 17 + i
        ws.write(rr, 0, row["Exposure Category"], fmt_text)
        ws.write_number(rr, 1, int(row["Customers"]), fmt_number)
        ws.write_number(rr, 2, float(row["Outstanding Amount"]), fmt_money)
        ws.write_number(rr, 3, float(row["% of Direction Total"]), fmt_percent)

    for i, row in due_summary.iterrows():
        rr = 17 + i
        ws.write(rr, 8, row["Exposure Category"], fmt_text)
        ws.write_number(rr, 9, int(row["Customers"]), fmt_number)
        ws.write_number(rr, 10, float(row["Outstanding Amount"]), fmt_money)
        ws.write_number(rr, 11, float(row["% of Direction Total"]), fmt_percent)

    # -------------------------
    # Top 15 tables
    # -------------------------
    ws.merge_range("A23:H23", "TOP CUSTOMERS OWING US — RANKED BY OUTSTANDING", fmt_section)
    ws.merge_range("J23:P23", "TOP CUSTOMERS WE OWE TOKEN — RANKED BY TOKEN DUE", fmt_section)

    top_headers = ["Rank", "Customer", "VPS Purchases", "Frequency", "Outstanding (₦)", "% of Total", "Exposure Category", "KYC Status"]
    for c, h in enumerate(top_headers):
        ws.write(23, c, h, fmt_header)

    due_top_headers = ["Rank", "Customer", "VPS Purchases", "Frequency", "Token Due (₦)", "% of Total", "Exposure Category"]
    for c, h in enumerate(due_top_headers, start=9):
        ws.write(23, c, h, fmt_header)

    def category_format(category):
        if category == "Critical":
            return fmt_critical
        if category == "High Materiality":
            return fmt_high
        if category == "Significant":
            return fmt_significant
        return fmt_low

    for i, (_, r) in enumerate(recovery.head(15).iterrows(), start=24):
        ws.write_number(i, 0, int(r["Rank"]), fmt_number)
        ws.write(i, 1, clean_id(r["Customer Name"]), fmt_text)
        ws.write_number(i, 2, int(r["Total VPS Payment Count"]), fmt_number)
        ws.write(i, 3, r["Purchase Frequency Rating"], fmt_text)
        ws.write_number(i, 4, float(r["Outstanding Amount"]), fmt_money)
        ws.write_number(i, 5, float(r["Individual Share of Total"]), fmt_percent)
        ws.write(i, 6, r["Exposure Category"], category_format(r["Exposure Category"]))
        ws.write(i, 7, r["KYC Status"], fmt_text)

    for i, (_, r) in enumerate(token_due.head(15).iterrows(), start=24):
        ws.write_number(i, 9, int(r["Rank"]), fmt_number)
        ws.write(i, 10, clean_id(r["Customer Name"]), fmt_text)
        ws.write_number(i, 11, int(r["Total VPS Payment Count"]), fmt_number)
        ws.write(i, 12, r["Purchase Frequency Rating"], fmt_text)
        ws.write_number(i, 13, float(r["Outstanding Amount"]), fmt_money)
        ws.write_number(i, 14, float(r["Individual Share of Total"]), fmt_percent)
        ws.write(i, 15, r["Exposure Category"], category_format(r["Exposure Category"]))

    # -------------------------
    # Full supporting detail
    # -------------------------
    detail_start = 41  # Excel row 42
    ws.merge_range(detail_start, 0, detail_start, 15, "FULL OPTION 2 CUSTOMER DETAIL", fmt_section)
    detail_header_row = detail_start + 1

    detail_cols = [
        "Direction",
        "Rank",
        "Summary Customer Key",
        "Customer Name",
        "Customer Classification",
        "KYC Status",
        "Total VPS Payment Count",
        "Purchase Frequency Rating",
        "VPS Paid / No Token Count",
        "Token / No VPS Payment Count",
        "Option2 Expected Token Value",
        "Option2 Actual Macron Token",
        "Outstanding Amount",
        "Individual Share of Total",
        "Exposure Category",
        "Relevant Total Outstanding",
    ]

    for c, h in enumerate(detail_cols):
        ws.write(detail_header_row, c, h, fmt_header)

    full_detail = option2_detail.sort_values(["Direction", "Rank"]).reset_index(drop=True)
    for ridx, r in full_detail.iterrows():
        rr = detail_header_row + 1 + ridx
        for cidx, col in enumerate(detail_cols):
            value = r.get(col)
            if pd.isna(value):
                ws.write_blank(rr, cidx, None, fmt_text)
            elif col in {"Option2 Expected Token Value", "Option2 Actual Macron Token", "Outstanding Amount", "Relevant Total Outstanding"}:
                ws.write_number(rr, cidx, float(value), fmt_money)
            elif col == "Individual Share of Total":
                ws.write_number(rr, cidx, float(value), fmt_percent)
            elif col in {"Rank", "Total VPS Payment Count", "VPS Paid / No Token Count", "Token / No VPS Payment Count"}:
                ws.write_number(rr, cidx, int(value), fmt_number)
            elif col == "Exposure Category":
                ws.write(rr, cidx, value, category_format(value))
            else:
                ws.write(rr, cidx, clean_id(value), fmt_text)

    # Filters only on the full supporting detail table.
    if len(full_detail):
        ws.autofilter(detail_header_row, 0, detail_header_row + len(full_detail), len(detail_cols) - 1)

    # -------------------------
    # Layout
    # -------------------------
    widths = [18, 9, 25, 30, 22, 24, 17, 18, 18, 18, 21, 21, 20, 14, 20, 22]
    for c, width in enumerate(widths):
        ws.set_column(c, c, width)

    ws.freeze_panes(detail_header_row + 1, 0)
    ws.hide_gridlines(2)


def build_management(final_recon, kyc, pay, control):
    paid_no_token = final_recon[(final_recon["VPS Present?"] == "YES") & (final_recon["Macron Present?"] == "NO")]
    token_no_vps = final_recon[(final_recon["Macron Present?"] == "YES") & (final_recon["VPS Present?"] == "NO")]
    missing_kyc = final_recon[final_recon["FINAL_KYC_ID"].map(clean_id).eq("")]
    metrics = [
        ("Full reconciliation rows", len(final_recon), "count"),
        ("Fully reconciled rows", int((final_recon["FINAL_Reconciliation_Status"] == "FULLY RECONCILED").sum()), "count"),
        ("KYC customers", len(kyc), "people"),
        ("Rows with no KYC", len(missing_kyc), "count"),
        ("VPS payments with no Macron token", len(paid_no_token), "count"),
        ("Customers with VPS payment but no Macron token", paid_no_token["FINAL_KYC_ID"].replace("", np.nan).nunique(), "people"),
        ("Customer amount paid on VPS but no token", pd.to_numeric(paid_no_token["VPS Customer Amount Paid"], errors="coerce").sum(), "amount"),
        ("Expected token value not issued (VPS paid - 100)", pd.to_numeric(paid_no_token["Expected Macron Token (VPS Paid - 100)"], errors="coerce").sum(), "amount"),
        ("Macron tokens with no VPS payment", len(token_no_vps), "count"),
        ("Customers with Macron token but no VPS payment", token_no_vps["FINAL_KYC_ID"].replace("", np.nan).nunique(), "people"),
        ("Macron token value with no VPS payment", pd.to_numeric(token_no_vps["Macron Token Amount"], errors="coerce").sum(), "amount"),
        ("Paymeter rows cleaned", len(pay), "count"),
        ("Paymeter rows with Address spill repaired", int(pay.attrs.get("repaired_rows", 0)), "count"),
        ("Paymeter Address spill cells deleted", int(pay.attrs.get("spill_cells_deleted", 0)), "count"),
        ("Customer Summary tie-out controls passed", int((control["Status"] == "MATCH").sum()), "count"),
        ("Customer Summary tie-out controls failed", int((control["Status"] != "MATCH").sum()), "count"),
    ]
    return pd.DataFrame(metrics, columns=["Metric", "Value", "Unit"])


def build_source_coverage(final_recon, bank, vps, pay, mac):
    checks = [
        ("Bank Transactions", len(bank), "LEFT_BANK_BANK_Row"),
        ("VPS", len(vps), "LEFT_VPS_VPS_Row"),
        ("Paymeter", len(pay), "RIGHT_PM_PM_Row"),
        ("Macron", len(mac), "RIGHT_MAC_MAC_Row"),
    ]
    rows = []
    for source, expected, col in checks:
        vals = pd.to_numeric(final_recon.get(col, pd.Series(dtype=float)), errors="coerce").dropna().astype(int)
        represented = len(vals)
        unique = vals.nunique()
        rows.append({
            "Source": source,
            "Expected Rows": expected,
            "Rows Represented": represented,
            "Unique Rows Represented": unique,
            "Duplicate Uses": represented - unique,
            "Missing Rows": expected - unique,
            "Status": "COMPLETE" if expected == unique and represented == unique else "REVIEW",
        })
    return pd.DataFrame(rows)


# ================================================================
# COLUMN GUIDE
# ================================================================

def build_column_guide(sheet_pairs):
    purpose = {
        "01 Cleaned Paymeter": "Paymeter after CSV repair/normalisation. RRN and meter scientific-notation loss is explicitly flagged.",
        "02 PM Macron Merge": "Supporting transaction history: Paymeter Reference = Macron REFERENCE ID.",
        "03 VPS Bank Merge": "VPS reconciled to Bank credits; all other genuine dated Bank credits/debits remain as Bank-only movements.",
        "04 KYC": "Master customer identity built from the uploaded KYC mapping; conflicts are never silently resolved.",
        "05 PM Macron + KYC": "Paymeter-Macron supporting history with uploaded KYC attached.",
        "06 VPS Bank + KYC": "VPS-Bank history with uploaded KYC attached.",
        "07 Full Reconciliation": "Main reconciliation. VPS is the control ledger and Macron is the token ledger; Paymeter/KYC provide identity and reference evidence.",
        "08 VPS vs Macron": "Focused VPS-versus-Macron view. Expected token = VPS transaction_amount_minor - 100.",
        "09 Customer Summary": "Customer/source summary with unlinked rows retained so source totals do not disappear.",
        "10 Summary Control": "Hard tie-out proving Customer Summary totals equal source totals.",
        "11 Reconciliation Summary": "Management exception and value summary.",
        "12 Paid No Token": "VPS payment rows with no paired Macron token.",
        "13 Token No VPS": "Macron token rows with no paired VPS payment.",
        "14 Missing KYC": "Full-reconciliation rows without an uploaded-KYC identity.",
        "15 Source Coverage": "Checks that every genuine Bank transaction, VPS, Paymeter and Macron source row is represented exactly once.",
        "16 Option 2 Management": "Management exposure report retained from the original application.",
        "17 Cash Posting": "Seven-column posting view. Unrelated Bank credits/debits remain with blank customer/account and actual Bank Cash In/Cash Out.",
        "18 Uploaded KYC": "Raw uploaded KYC file retained for traceability.",
        "19 VPS Paymeter Bridge": "Transaction-level VPS-to-Paymeter bridge using account, exact gross amount and nearest time.",
        "20 KYC Conflicts": "Uploaded KYC or observed Paymeter account mapping conflicts requiring review.",
    }
    rows = []
    for sheet, df in sheet_pairs:
        for col in df.columns:
            calc = "Copied from source / merged source row."
            meaning = col
            if col == "Expected Macron Token (VPS Paid - 100)":
                calc = "VPS transaction_amount_minor - 100."
                meaning = "Token value Macron should issue for the VPS payment."
            elif col == "Difference Expected Macron - Actual Macron":
                calc = "(VPS transaction_amount_minor - 100) - Macron AMOUNT."
                meaning = "Primary VPS-to-Macron value difference."
            elif col == "PMMAC_Match_Status":
                calc = "MATCHED when Paymeter Reference = Macron REFERENCE ID; this is supporting transaction evidence."
            elif col == "VPSBANK_Match_Status":
                calc = "MATCHED when VPS session_id is found in Bank Narration or unique settled amount/date fallback succeeds."
            elif col == "KYC_ID":
                calc = "Uploaded KYC PAYMETER - Account Number, which is also the VPS virtual account identity."
            elif col == "FINAL_Match_Method":
                calc = "Tier 1: VPS->Paymeter transaction + exact Paymeter/Macron reference. Tier 2: KYC Macron account + expected value + time. Tier 3: same KYC Macron account + nearest time, REVIEW only."
            elif col == "FINAL_Reconciliation_Status":
                calc = "Derived primarily from VPS/Macron presence, KYC identity, expected-token difference and Bank control."
            elif col == "Cash In":
                calc = "VPS gross - N100 fixed charge, matching the requested posting logic."
            elif col == "Cash Out":
                calc = "Actual Macron token amount."
            elif col == "Interest":
                calc = "N100 - actual bank charge; actual bank charge = VPS gross - VPS settled."
            elif col == "Source":
                calc = "Bank/Macron when both are present; otherwise the available source side."
            elif col == "Option2 Expected Token Value":
                calc = "Total VPS Customer Amount Paid - (Total VPS Payment Count x N100)."
                meaning = "Correct token value due to the customer after the N100 charge on each VPS payment."
            elif col == "Option2 Actual Macron Token":
                calc = "Total Macron Token Amount from Customer Summary."
                meaning = "Actual aggregate token value issued by Macron."
            elif col == "Outstanding Amount":
                calc = "Absolute directional difference between actual Macron token and expected token."
                meaning = "Amount recoverable from customer or token value still due to customer."
            elif col == "Individual Share of Total":
                calc = "Customer Outstanding Amount divided by total outstanding for that Direction."
                meaning = "Option 2 financial materiality measure."
            elif col == "Exposure Category":
                calc = "Critical >=10%; High Materiality 2%-<10%; Significant 0.5%-<2%; Low Materiality <0.5%."
            elif col == "Purchase Frequency Rating":
                calc = "Very Frequent 13+; Frequent 7-12; Regular 4-6; Occasional 1-3; No VPS History 0."
            elif col.startswith("Difference"):
                calc = "Arithmetic difference between the named values. Zero means exact alignment."
            elif col.startswith("Total "):
                calc = "Sum/count of the named value for the customer/source identity group."
            elif col.startswith("PM_"):
                calc = "Paymeter field after cleaning, prefixed PM_."
            elif col.startswith("MAC_"):
                calc = "Macron source field, prefixed MAC_."
            elif col.startswith("VPS_"):
                calc = "VPS source field, prefixed VPS_."
            elif col.startswith("BANK_"):
                calc = "Bank source field, prefixed BANK_."
            rows.append({"Sheet": sheet, "Sheet Purpose": purpose.get(sheet, "Reconciliation output."), "Column Header": col, "Meaning": meaning, "How Calculated / Arrived At": calc})
    return pd.DataFrame(rows)

# ================================================================
# EXCEL OUTPUT
# ================================================================

def make_excel(result):
    output = io.BytesIO()
    sheets = [
        ("01 Cleaned Paymeter", result["pay_cleaned"]),
        ("02 PM Macron Merge", result["pm_mac_merge"]),
        ("03 VPS Bank Merge", result["vps_bank_merge"]),
        ("04 KYC", result["kyc"]),
        ("05 PM Macron + KYC", result["pm_mac_kyc"]),
        ("06 VPS Bank + KYC", result["vps_bank_kyc"]),
        ("07 Full Reconciliation", result["final_recon"]),
        ("08 VPS vs Macron", result["vps_macron"]),
        ("09 Customer Summary", result["customer_summary"]),
        ("10 Summary Control", result["summary_control"]),
        ("11 Reconciliation Summary", result["management"]),
        ("12 Paid No Token", result["paid_no_token"]),
        ("13 Token No VPS", result["token_no_vps"]),
        ("14 Missing KYC", result["missing_kyc"]),
        ("15 Source Coverage", result["source_coverage"]),
        ("17 Cash Posting", result["cash_posting"]),
        ("18 Uploaded KYC", result["kyc_uploaded"]),
        ("19 VPS Paymeter Bridge", result["vps_pay_bridge"]),
        ("20 KYC Conflicts", result["kyc_conflicts"]),
    ]
    guide_pairs = sheets + [("16 Option 2 Management", result["option2_management"])]
    guide = build_column_guide(guide_pairs)
    sheets.insert(0, ("00 Column Guide", guide))

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd-mmm-yyyy hh:mm:ss", engine_kwargs={"options": {"strings_to_urls": False, "strings_to_formulas": False}}) as writer:
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "text_wrap": True, "valign": "vcenter"})
        fmt_money = wb.add_format({"num_format": '#,##0.00;[Red](#,##0.00);-'})
        fmt_date = wb.add_format({"num_format": 'dd-mmm-yyyy'})
        fmt_wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        fmt_good = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        fmt_bad = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        for name, df in sheets:
            safe = df.drop(columns=[c for c in df.columns if c.endswith("_Date") or c.endswith("_DateTime")], errors="ignore").copy()
            safe.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            ws.freeze_panes(1, 0)
            if len(safe.columns):
                ws.autofilter(0, 0, max(len(safe), 1), len(safe.columns) - 1)
            for j, c in enumerate(safe.columns):
                ws.write(0, j, c, fmt_header)
                lower = c.lower()
                width, cell_fmt = 16, None
                if c == "Transaction Date":
                    width, cell_fmt = 16, fmt_date
                elif any(x in lower for x in ["name", "narration", "method", "status", "address", "purpose", "meaning", "calculated", "notes", "conflict"]):
                    width, cell_fmt = 30, fmt_wrap
                elif any(x in lower for x in ["amount", "credit", "charge", "difference", "total", "cash in", "cash out", "interest"]):
                    width, cell_fmt = 18, fmt_money
                else:
                    width = min(max(len(c) + 2, 12), 24)
                ws.set_column(j, j, width, cell_fmt)
            for status_col in ["Status", "FINAL_Reconciliation_Status", "FINAL_Amount_Status", "FINAL_Supporting_Check_Status", "VPS_vs_Macron_Status", "PMMAC_Match_Status", "VPSBANK_Match_Status", "Bridge_Match_Status", "KYC Status"]:
                if status_col in safe.columns and len(safe):
                    j = safe.columns.get_loc(status_col)
                    ws.conditional_format(1, j, len(safe), j, {"type": "text", "criteria": "containing", "value": "MATCH", "format": fmt_good})
                    for bad_word in ["NO ", "DIFFERENCE", "CONFLICT", "REVIEW", "AMBIGUOUS"]:
                        ws.conditional_format(1, j, len(safe), j, {"type": "text", "criteria": "containing", "value": bad_word, "format": fmt_bad})

        _write_option2_management_sheet(writer, result["option2_management"])

    output.seek(0)
    return output

# ================================================================
# MASTER ENGINE — EXACT USER-REQUESTED ORDER
# ================================================================

def run_reconciliation(bank_file, vps_file, paymeter_file, macron_file, kyc_file,
                       date_tolerance=DATE_TOLERANCE_DAYS,
                       transaction_hours=TRANSACTION_TIME_TOLERANCE_HOURS,
                       progress=None):
    progress_call(progress, 4, "Reading and cleaning all five source files...")
    pay = clean_paymeter(paymeter_file)
    mac = read_macron(macron_file)
    vps = read_vps(vps_file)
    bank = read_bank(bank_file)
    kyc_uploaded = read_kyc(kyc_file)

    progress_call(progress, 14, "Building supporting Paymeter-to-Macron reference history...")
    pm_mac = merge_paymeter_macron(pay, mac)

    progress_call(progress, 25, "Reconciling VPS receipts to the Bank statement...")
    vps_bank = merge_vps_bank(vps, bank, date_tolerance)

    progress_call(progress, 36, "Building uploaded-KYC master and checking mapping conflicts...")
    kyc = build_kyc(kyc_uploaded, pay, mac, vps_bank)
    kyc_conflicts = kyc[kyc["KYC Status"].str.contains("CONFLICT", case=False, na=False)].copy()

    progress_call(progress, 46, "Attaching KYC to supporting Paymeter/Macron and VPS/Bank histories...")
    pm_mac_kyc = attach_kyc_to_pm_mac(pm_mac, kyc)
    vps_bank_kyc = attach_kyc_to_vps_bank(vps_bank, kyc)

    progress_call(progress, 57, "Bridging VPS transactions to Paymeter by account, gross value and time...")
    vps_pay_bridge = build_vps_paymeter_bridge(vps_bank_kyc, pay, transaction_hours, HIGH_CONFIDENCE_MINUTES)

    progress_call(progress, 70, "Reconciling VPS directly against Macron using reference, KYC, value and time...")
    final_recon = reconcile_merged_reports(vps_bank_kyc, vps_pay_bridge, pay, mac, kyc, transaction_hours, HIGH_CONFIDENCE_MINUTES)

    progress_call(progress, 81, "Building VPS-versus-Macron and seven-column Cash In/Cash Out views, including Bank-only movements...")
    vps_macron = build_vps_macron_comparison(final_recon)
    cash_posting = build_posting_ledger(final_recon)

    progress_call(progress, 89, "Building complete Customer Summary and hard source tie-out...")
    customer_summary = build_customer_summary(final_recon, kyc)
    summary_control = build_summary_control(customer_summary, bank, vps, pay, mac)
    validate_tieout(summary_control)

    paid_no_token = final_recon[(final_recon["VPS Present?"] == "YES") & (final_recon["Macron Present?"] == "NO")].copy()
    token_no_vps = final_recon[(final_recon["Macron Present?"] == "YES") & (final_recon["VPS Present?"] == "NO")].copy()
    missing_kyc = final_recon[final_recon["FINAL_KYC_ID"].map(clean_id).eq("")].copy()
    source_coverage = build_source_coverage(final_recon, bank, vps, pay, mac)
    management = build_management(final_recon, kyc, pay, summary_control)

    progress_call(progress, 96, "Building retained Option 2 management exposure report...")
    option2_management = build_option2_management(customer_summary)

    progress_call(progress, 100, "Reconciliation calculations complete.")
    return {
        "pay_cleaned": pay,
        "pm_mac_merge": pm_mac,
        "vps_bank_merge": vps_bank,
        "kyc_uploaded": kyc_uploaded,
        "kyc": kyc,
        "kyc_conflicts": kyc_conflicts,
        "pm_mac_kyc": pm_mac_kyc,
        "vps_bank_kyc": vps_bank_kyc,
        "vps_pay_bridge": vps_pay_bridge,
        "final_recon": final_recon,
        "vps_macron": vps_macron,
        "cash_posting": cash_posting,
        "customer_summary": customer_summary,
        "summary_control": summary_control,
        "management": management,
        "paid_no_token": paid_no_token,
        "token_no_vps": token_no_vps,
        "missing_kyc": missing_kyc,
        "source_coverage": source_coverage,
        "option2_management": option2_management,
    }

# ================================================================
# STREAMLIT
# ================================================================

def main():
    if st is None:
        raise RuntimeError("Install Streamlit and the packages listed in requirements.txt")
    st.set_page_config(page_title="KYC-Driven VPS / Macron Reconciliation", page_icon="📊", layout="wide")
    st.title("📊 KYC-Driven VPS / Macron Reconciliation")
    st.caption("Bank → VPS → Paymeter bridge/KYC identity → Macron. VPS and Macron are the primary reconciliation ledgers.")
    st.info(
        "Uploaded KYC is the master identity table. Paymeter supplies the account bridge and transaction reference evidence. "
        "Expected Macron token is VPS transaction_amount_minor less ₦100. Ambiguous matches are not forced."
    )

    with st.sidebar:
        bank_file = st.file_uploader("Providus Bank Statement", type=["xlsx", "xls"])
        vps_file = st.file_uploader("VPS Report", type=["xlsx", "xls"])
        pay_file = st.file_uploader("Paymeter Report", type=["csv", "xlsx", "xls"])
        mac_file = st.file_uploader("Macron Report", type=["xlsx", "xls"])
        kyc_file = st.file_uploader("KYC Master / Account Mapping", type=["xlsx", "xls"])
        date_tol = st.slider("Bank fallback date tolerance (days)", 0, 7, DATE_TOLERANCE_DAYS)
        transaction_hours = st.slider("Transaction matching window (hours)", 1, 72, TRANSACTION_TIME_TOLERANCE_HOURS)
        st.metric("Fixed customer charge", "₦100.00")
        run = st.button("RUN RECONCILIATION", type="primary", use_container_width=True)

    if not run:
        st.markdown("""
### Processing order
1. Clean Paymeter and flag identifiers whose precision was lost in export.
2. Reconcile VPS to the actual Bank statement.
3. Load uploaded KYC as the master mapping: **VPS virtual account / Paymeter account → Macron account**.
4. Retain Paymeter ↔ Macron exact reference history as supporting evidence.
5. Bridge each VPS transaction to Paymeter using account number, exact gross value and nearest time.
6. Reconcile **VPS directly against Macron**:
   - exact Paymeter Reference = Macron REFERENCE ID first;
   - then KYC Macron account + expected value + time;
   - then same KYC Macron account + nearest time as REVIEW-only fallback.
7. Preserve every unmatched Bank, VPS, Paymeter and Macron source row and hard-tie the customer summary back to source totals.
8. Produce the retained management report plus the seven-column **Cash Posting** sheet.
""")
        return

    if not all([bank_file, vps_file, pay_file, mac_file, kyc_file]):
        st.error("Upload all five reports: Bank, VPS, Paymeter, Macron and KYC.")
        return

    bar = st.progress(0)
    stage = st.empty()
    def cb(pct, msg):
        bar.progress(int(pct))
        stage.info(msg)

    try:
        result = run_reconciliation(bank_file, vps_file, pay_file, mac_file, kyc_file, date_tol, transaction_hours, cb)
        stage.success("Reconciliation calculation completed. Building Excel workbook...")
        excel = make_excel(result)
        stage.success("Excel workbook ready.")

        st.subheader("Reconciliation Summary")
        st.dataframe(result["management"], use_container_width=True, hide_index=True)
        tabs = st.tabs([
            "Full Reconciliation",
            "VPS vs Macron",
            "Cash Posting",
            "KYC",
            "KYC Conflicts",
            "Customer Summary",
            "Summary Control",
            "Source Coverage",
            "Option 2 Management",
        ])
        with tabs[0]:
            st.dataframe(result["final_recon"].head(1500), use_container_width=True, hide_index=True)
        with tabs[1]:
            st.dataframe(result["vps_macron"].head(1500), use_container_width=True, hide_index=True)
        with tabs[2]:
            st.dataframe(result["cash_posting"].head(1500), use_container_width=True, hide_index=True)
        with tabs[3]:
            st.dataframe(result["kyc"].head(2000), use_container_width=True, hide_index=True)
        with tabs[4]:
            st.dataframe(result["kyc_conflicts"].head(500), use_container_width=True, hide_index=True)
        with tabs[5]:
            st.dataframe(result["customer_summary"].head(1500), use_container_width=True, hide_index=True)
        with tabs[6]:
            st.dataframe(result["summary_control"], use_container_width=True, hide_index=True)
        with tabs[7]:
            st.dataframe(result["source_coverage"], use_container_width=True, hide_index=True)
        with tabs[8]:
            opt2 = result["option2_management"].copy()
            rec = opt2[opt2["Direction"] == "Customer Owes Us"]
            due = opt2[opt2["Direction"] == "We Owe Customer Token"]
            rec_total = pd.to_numeric(rec["Outstanding Amount"], errors="coerce").fillna(0).sum()
            due_total = pd.to_numeric(due["Outstanding Amount"], errors="coerce").fillna(0).sum()
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Customers owing us", len(rec))
            c2.metric("Recovery outstanding", f"₦{rec_total:,.2f}")
            c3.metric("Customers we owe token", len(due))
            c4.metric("Token value due", f"₦{due_total:,.2f}")
            c5.metric("Net receivable", f"₦{(rec_total - due_total):,.2f}")
            st.caption("Exposure Category remains based on each customer's share of the relevant outstanding total; purchase frequency is separate.")
            st.markdown("#### Customers owing us")
            st.dataframe(rec.sort_values("Rank").head(150), use_container_width=True, hide_index=True)
            st.markdown("#### Customers we owe token")
            st.dataframe(due.sort_values("Rank").head(150), use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ DOWNLOAD RECONCILIATION WORKBOOK",
            data=excel,
            file_name="KYC_Driven_VPS_Macron_Reconciliation_V11.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    except Exception as exc:
        stage.empty()
        st.exception(exc)


if __name__ == "__main__":
    main()
