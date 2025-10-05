import pandas as pd
import io
import streamlit as st
from typing import Tuple, List, Dict, Set

# ✅ Exact required column sequence (strict)
COLUMN_SEQUENCE: List[str] = [
    "CUST_ID", "CUST_NAME", "QUEUE", "OFC", "HOME", "MOBILE_NO", "TU", "ADDRESS",
    "EMAIL", "GENDER", "TCL", "OB", "BOS", "AOD", "MAD", "PDA", "LAST PAYMENT AMOUNT",
    "LAST_PAYMENT_DATE", "RESPONSE_CODE", "LAST_CONTACT_DATE", "BEHAVIOURAL_SCORE",
    "DELINQUENCY_STRING", "ADA_ACCOUNT", "DEBIT_AMOUNT_PREFERENCE", "MS", "DOSRI_FLAG",
    "EMPLOYEE_CODE", "RM_NUMBER", "COLLECTION_CYCLE", "UNIT_CODE", "UNIT_DESC",
    "LAST_ACTION_CODE", "RISK", "AGING", "HO FLAG", "BIRTHDATE", "UNIBANKER",
    "NS with tranx", "TPAP", "CRISPR", "block_code", "MEMO_LINE", "D_CUST_OPN",
    "AREA CODE", "PTP DATE", "CATEGORY/CLASSIF", "LAST DUE DATE", "Balance Type",
    "HO AMOUNT", "PRIO LIST", "PTP AMOUNT", "CONTACTED BY", "CLASSIFICATION",
    "TPAP DD", "AGENCY", "CLASSIF 2", "INHOUSE", "EMAIL NOTI", "CATEGORY",
    "SPOUSE NUMBER", "PTP FROM", "ADDRESS_1", "ADDRESS_2", "ADDRESS_3",
    "ADEPTRA RESULT", "USER_FLG8", "P_RESON_CD", "TP_PDR_Code", "EMPLOYMENT",
    "EXCLUSION", "PREDEL NOTIF", "PUSHBACK STAT", "ACQUISITION CHANNEL",
    "OCCUPATION", "TYPE"
]

ACCEPTED_COLUMNS: Set[str] = set(COLUMN_SEQUENCE)


# -----------------------
# 📘 File Reading
# -----------------------
def read_excel_file(uploaded_file) -> Tuple[pd.DataFrame, List[str]]:
    """Read Excel file and return DataFrame and column headers."""
    try:
        df = pd.read_excel(uploaded_file)
        headers = [str(col).strip() for col in df.columns]
        df.columns = headers
        return df, headers
    except Exception as e:
        st.error(f"Error reading file {getattr(uploaded_file, 'name', '')}: {str(e)}")
        return None, []


# -----------------------
# ⚙️ Auto Header Alignment
# -----------------------
def align_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Automatically align headers to the master column sequence."""
    df_cols = [str(c).strip() for c in df.columns]

    # Add missing columns as None
    for col in COLUMN_SEQUENCE:
        if col not in df_cols:
            df[col] = None

    # Keep columns in exact sequence
    aligned_df = df.reindex(columns=COLUMN_SEQUENCE)
    return aligned_df


# -----------------------
# 📂 File Consolidation
# -----------------------
def consolidate_files(uploaded_files) -> Tuple[pd.DataFrame, Dict]:
    """Combine multiple Excel files into one, auto-aligning their columns."""
    if not uploaded_files:
        return None, {}

    stats = {
        "total_files": len(uploaded_files),
        "processed_files": 0,
        "skipped_files": [],
        "total_rows": 0
    }

    all_dfs = []

    for file in uploaded_files:
        df, headers = read_excel_file(file)
        if df is None:
            stats["skipped_files"].append(file.name)
            continue

        # Auto-align headers
        aligned_df = align_headers(df)

        # Add aligned DataFrame to list
        all_dfs.append(aligned_df)
        stats["processed_files"] += 1
        stats["total_rows"] += len(aligned_df)

    if not all_dfs:
        st.error("❌ No valid files to consolidate.")
        return None, stats

    consolidated_df = pd.concat(all_dfs, ignore_index=True)
    consolidated_df = consolidated_df.reindex(columns=COLUMN_SEQUENCE)

    return consolidated_df, stats


# -----------------------
# 💾 Excel Download
# -----------------------
def generate_excel_download(df: pd.DataFrame, filename: str) -> bytes:
    """Return Excel bytes for download."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
    return output.getvalue()
