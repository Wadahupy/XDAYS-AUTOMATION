import pandas as pd
import re
import io
import streamlit as st
from typing import Tuple, List, Dict, Set
from datetime import datetime
import msoffcrypto
from io import BytesIO
import re
import numpy as np

# Configure pandas to use new behavior for type conversion
pd.set_option('future.no_silent_downcasting', True)

# ----------------------------
# HEADER MAPPING (with alternates)
# ----------------------------
HEADER_MAPPING = {
    # Standard headers
    "CUST_ID": "CUST_ID",
    "CUST_NAME": "CUST_NAME",
    "UNIT_CODE": "UNIT_CODE",
    "OFFICE_PH": "OFC",
    "HOME_PH": "HOME",
    "MOBILE_NO": "MOBILE_NO",
    "TU_RESULT": "TU",
    "EMAIL": "EMAIL",
    "AMOUNT_OUTSTANDING": "OB",  # ✅
    "XDAYS": "PDA",
    "LAST_PAYMENT_AMT": "LAST PAYMENT AMOUNT",
    "LAST_PAYMENT_DATE": "LAST_PAYMENT_DATE",
    "RISK": "RISK",
    "AGING": "AGING",
    "HO_FLAG": "HO_FLAG",
    "BIRTHDATE": "BIRTHDATE",
    "PTP_AMOUNT": "PTP AMOUNT",
    "PTP_DATE": "PTP_DATE",
    "TPAP_DD": "TPAP DD",
    "UNIBANKER": "UNIBANKER",
    "CLASSIFICATION": "CLASSIFICATION",
    "LAST_DUE_DATE (manual)": "LAST DUE DATE",
    "Balance_Type (100k up, etc)": "Balance Type",
    "BLOCK_CODE": "BLOCK_CODE",
    "BLOCK CODE": "BLOCK_CODE",
    "MEMO_LINE": "MEMO_LINE",
    "INHOUSE": "INHOUSE",
    "D_CUST_OPN": "D_CUST_OPN",  # ✅ Ensure captured when directly named this way

    
    # Alternate header names for D_CUST_OPN
    "DATE_CUST_OPENED": "D_CUST_OPN",
    "CUST_OPEN_DATE": "D_CUST_OPN",
    "DATE_OPENED": "D_CUST_OPN",
    "OPEN_DATE": "D_CUST_OPN",
    "D CUST OPEN": "D_CUST_OPN",

    # Alternate header names for other financial columns
    "BOS/CB": "BOS",
    "BOS_CB": "BOS",
    "BOSCB": "BOS",
    "AMOUNT_OVERDUE/AOD": "AOD",
    "AMOUNT_OVERDUE": "AOD",
    "AMT_OVERDUE": "AOD",
    "MIN_AMOUNT_DUE/MAD": "MAD",
    "MIN_AMOUNT_DUE": "MAD",
    "MIN_AMT_DUE": "MAD",
}



# ----------------------------
# MASTER COLUMN SEQUENCE
# ----------------------------
COLUMN_SEQUENCE: List[str] = [
    "CUST_ID", "CUST_NAME", "QUEUE", "OFC", "HOME", "MOBILE_NO", "TU", "ADDRESS",
    "EMAIL", "GENDER", "TCL", "OB", "BOS", "AOD", "MAD", "PDA", "LAST PAYMENT AMOUNT",
    "LAST_PAYMENT_DATE", "RESPONSE_CODE", "LAST_CONTACT_DATE", "BEHAVIOURAL_SCORE",
    "DELINQUENCY_STRING", "ADA_ACCOUNT", "DEBIT_AMOUNT_PREFERENCE", "MS", "DOSRI_FLAG",
    "EMPLOYEE_CODE", "RM_NUMBER", "COLLECTION_CYCLE", "UNIT_CODE", "UNIT_DESC",
    "LAST_ACTION_CODE", "RISK", "AGING", "HO_FLAG", "BIRTHDATE", "UNIBANKER",
    "NS_with_tranx", "TPAP", "CRISPR", "BLOCK_CODE", "MEMO_LINE", "D_CUST_OPN",
    "AREA CODE", "PTP DATE", "CATEGORY/CLASSIF", "LAST DUE DATE", "Balance Type",
    "HO AMOUNT", "PRIO LIST", "PTP AMOUNT", "CONTACTED BY", "CLASSIFICATION",
    "TPAP DD", "AGENCY", "CLASSIF 2", "INHOUSE", "EMAIL NOTI", "CATEGORY",
    "SPOUSE NUMBER", "PTP FROM", "ADDRESS_1", "ADDRESS_2", "ADDRESS_3",
    "ADEPTRA RESULT", "USER_FLG8", "P_RESON_CD", "TP_PDR_Code", "EMPLOYMENT",
    "EXCLUSION", "PREDEL NOTIF", "PUSHBACK STAT", "ACQUISITION CHANNEL",
    "OCCUPATION", "TYPE"
]

ACCEPTED_COLUMNS: Set[str] = set(COLUMN_SEQUENCE)


# ----------------------------
# 📘 READ EXCEL FILE
# ----------------------------


def read_excel_file(uploaded_file) -> Tuple[pd.DataFrame, List[str]]:
    """Read (and decrypt if needed) Excel file, ask user for password if encrypted, and return DataFrame and headers."""
    try:
        file_name = getattr(uploaded_file, "name", "")
        file_ext = file_name.split(".")[-1].lower()

        decrypted = BytesIO()

        # 🧩 Try opening with msoffcrypto to detect encryption
        try:
            office_file = msoffcrypto.OfficeFile(uploaded_file)
            # Check if encrypted
            if office_file.is_encrypted():
                st.warning(f"🔒 File '{file_name}' is password-protected.")
                manual_password = st.text_input(f"Enter password for {file_name}:", type="password", key=file_name)

                if not manual_password:
                    st.stop()  # Wait for user input before proceeding

                try:
                    office_file.load_key(password=manual_password)
                    office_file.decrypt(decrypted)
                    st.success(f"✅ Successfully decrypted {file_name}.")
                except Exception:
                    st.error("❌ Incorrect password. Please try again.")
                    st.stop()
            else:
                # Not encrypted
                uploaded_file.seek(0)
                decrypted = uploaded_file.read()
                decrypted = BytesIO(decrypted)
        except Exception:
            # Fallback: not encrypted or invalid format
            uploaded_file.seek(0)
            decrypted = uploaded_file.read()
            decrypted = BytesIO(decrypted)

        # � Choose engine automatically
        engine = "xlrd" if file_ext == "xls" else "openpyxl"
        decrypted.seek(0)
        df = pd.read_excel(decrypted, engine=engine)

        # 🧹 Clean header names
        df.columns = [str(c).strip() for c in df.columns]

        # Convert all object/string columns to string dtype immediately after reading
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).replace(['nan', 'None', 'NaT', 'null'], np.nan)
            if pd.notna(df[col]).any():  # Only convert if column has non-NA values
                df[col] = df[col].astype('string')

        return df, list(df.columns)

    except Exception as e:
        st.error(f"❌ Error reading file {file_name}: {str(e)}")
        return None, []


# ----------------------------
# 🧹 DATA CLEANING
# ----------------------------

def clean_data(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """
    Cleans and standardizes key columns based on the user's specifications.
    - Adds leading zeros to CUST_ID if missing (simulating =0&CUST_ID in Excel).
    - Fixes MOBILE_NO format.
    - Removes invalid values like '????', '00', '0', or errors.
    - Standardizes date columns to mm/dd/yyyy.
    - Adds COLLECTION_CYCLE column (detected from filename, e.g., c27 → 27).
    """
    df = df.copy()


    # =====================
    # 🔍 Extract collection cycle from filename
    # =====================
    match = re.search(r"c(\d+)", filename.lower())
    cycle = match.group(1) if match else "N/A"
    # Add leading zero if it's a single digit (e.g., "5" → "05")
    if cycle.isdigit() and len(cycle) == 1:
        cycle = f"0{cycle}"

    df["COLLECTION_CYCLE"] = cycle

    # =====================
    # 🧹 Clean CUST_ID
    # =====================
    if "CUST_ID" in df.columns:
        df["CUST_ID"] = df["CUST_ID"].astype(str).str.strip()
        df["CUST_ID"] = df["CUST_ID"].replace(["nan", "None", "NaT", "", "??", "????"], np.nan)

        # Create temporary column simulating Excel formula =0&CUST_ID
        df["TEMP_CUST_ID"] = df["CUST_ID"].apply(lambda x: f"0{x}" if pd.notna(x) and not str(x).startswith("0") else x)

        # Handle scientific or float-like values (e.g., 2.01002E+12 or 201002201707538.0)
        df["TEMP_CUST_ID"] = df["TEMP_CUST_ID"].apply(
            lambda x: str(int(float(x))) if isinstance(x, str) and re.match(r"^\d+\.0+$", x) else x
        )

        # Copy values back to main CUST_ID
        df["CUST_ID"] = df["TEMP_CUST_ID"]

        # Drop the temporary column (like deleting the Excel helper column)
        df.drop(columns=["TEMP_CUST_ID"], inplace=True)

    # =====================
    # 📱 Clean MOBILE_NO
    # =====================
    if "MOBILE_NO" in df.columns:
        df["MOBILE_NO"] = df["MOBILE_NO"].astype(str).str.strip()

        def normalize_mobile(num):
            if pd.isna(num) or str(num) in ['nan', 'None', '', 'NaT', '????', '00', '0']:
                return np.nan
            num = re.sub(r"\D", "", str(num))
            if not num:
                return np.nan
            if num.startswith("639") and len(num) == 12:
                return "0" + num[2:]
            elif len(num) == 10:
                return "0" + num
            elif len(num) == 11 and num.startswith("09"):
                return num
            return np.nan

        df["MOBILE_NO"] = df["MOBILE_NO"].apply(normalize_mobile)

    # =====================
    # 📅 Clean date columns
    # =====================
    date_column_mapping = {
        "LAST_PAYMENT_DATE": "LAST_PAYMENT_DATE",
        "LAST_CONTACT_DATE": "LAST_CONTACT_DATE",
        "PTP DATE": "PTP DATE",
        "PTP_DATE": "PTP DATE",
        "BIRTHDATE": "BIRTHDATE",
        "LAST DUE DATE": "LAST DUE DATE",
        "LAST_DUE_DATE": "LAST DUE DATE",
        "TPAP DD": "TPAP DD",
        "TPAP_DD": "TPAP DD",
        "D_CUST_OPN": "D CUST OPEN"
    }

    for col in df.columns:
        if col in date_column_mapping or col in date_column_mapping.values():
            df[col] = pd.to_datetime(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: x.strftime("%m/%d/%Y") if pd.notna(x) else np.nan)

    # =====================
    # 🔢 Numeric columns
    # =====================
    NUMERIC_COLUMNS = ["OB", "BOS", "AOD", "MAD", "PDA", "LAST PAYMENT AMOUNT",
                       "PTP AMOUNT", "HO AMOUNT", "TCL"]

    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"[^\d\.\-]", "", regex=True)
            df[col] = df[col].replace(["", "nan", "None", "NaT"], np.nan)
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # =====================
    # 🔤 String columns
    # =====================
    # Define all string columns including both variations of headers
    STRING_COLUMNS = [
        "MOBILE_NO", "CUST_ID", "EMAIL", "COLLECTION_CYCLE", "UNIT_CODE", 
        "EMPLOYEE_CODE", "GENDER", "RISK", "TYPE", "CATEGORY", 
        "CLASSIFICATION", "TU", "ADDRESS", "QUEUE", "UNIT_DESC", 
        "BLOCK_CODE", "BLOCK CODE", "MEMO_LINE", "UNIBANKER", "AGING", 
        "BALANCE_TYPE", "INHOUSE", "CATEGORY_CLASSIF", "SPOUSE_NUMBER", 
        "CUST_NAME", "RM_NUMBER", "LAST_ACTION_CODE", "HO_FLAG", "AGENCY", 
        "AREA_CODE", "CONTACTED_BY", "CLASSIF_2", "OFC", "HOME", 
        "OFFICE_PH", "HOME_PH"
    ]

    for col in STRING_COLUMNS:
        if col in df.columns:
            # Handle invalid values consistently
            invalid_values = ["nan", "None", "NaT", "????", "??"]
            if col in ["OFC", "HOME", "OFFICE_PH", "HOME_PH"]:  # Phone number columns
                invalid_values.extend(["0", "00"])
            
            # Convert to string, handle invalid values, and ensure consistent type
            df[col] = df[col].astype(str)
            df[col] = df[col].apply(lambda x: np.nan if str(x).strip() in invalid_values else str(x).strip())
            if pd.notna(df[col]).any():  # Only convert if column has non-NA values
                df[col] = df[col].astype('string')

    return df


# ----------------------------
# ⚙️ ALIGN HEADERS
# ----------------------------
def align_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Align bank headers to internal headers with normalization (handles spaces, case, underscores)."""

    def normalize(header: str) -> str:
        return re.sub(r'[\s_]+', '', str(header).strip().upper())

    # Create a normalized lookup dict for HEADER_MAPPING
    normalized_mapping = {normalize(k): v for k, v in HEADER_MAPPING.items()}

    new_cols = []
    for col in df.columns:
        norm_col = normalize(col)
        mapped_col = normalized_mapping.get(norm_col, col.strip())
        new_cols.append(mapped_col)

    df.columns = new_cols

    # Add missing columns from master template
    for col in COLUMN_SEQUENCE:
        if col not in df.columns:
            df[col] = np.nan

    # Reorder columns according to master
    return df.reindex(columns=COLUMN_SEQUENCE)



# ----------------------------
# 📂 CONSOLIDATE FILES
# ----------------------------
def consolidate_files(uploaded_files) -> Tuple[pd.DataFrame, Dict]:
    if not uploaded_files:
        return None, {}
    

    stats = {"total_files": len(uploaded_files), "processed_files": 0, "skipped_files": [], "total_rows": 0}
    all_dfs = []

    for file in uploaded_files:
        df, _ = read_excel_file(file)
        if df is None:
            stats["skipped_files"].append(file.name)
            continue

        df = clean_data(df, file.name)
        df = align_headers(df)

        all_dfs.append(df)
        stats["processed_files"] += 1
        stats["total_rows"] += len(df)

    if not all_dfs:
        st.error("❌ No valid files to consolidate.")
        return None, stats

    consolidated = pd.concat(all_dfs, ignore_index=True)
    consolidated = consolidated.reindex(columns=COLUMN_SEQUENCE)
    return consolidated, stats


# ----------------------------
# 💾 DOWNLOAD & DISPLAY
# ----------------------------
def prepare_dataframe_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare DataFrame for display in Streamlit by ensuring Arrow-compatible types."""
    display_df = df.copy()
    
    # Convert all columns to Arrow-compatible types
    for col in display_df.columns:
        if display_df[col].dtype == 'object':
            # Convert object columns to string, preserving NaN
            display_df[col] = display_df[col].apply(
                lambda x: str(x) if pd.notna(x) and str(x) not in ['nan', 'None', ''] else None
            )
        elif pd.api.types.is_numeric_dtype(display_df[col]):
            # Ensure numeric columns are proper float/int
            display_df[col] = pd.to_numeric(display_df[col], errors='coerce')
    
    return display_df

def generate_excel_download(df: pd.DataFrame, filename: str) -> bytes:
    """Generate Excel file for download, ensuring clean data output."""
    # Create a copy for export
    export_df = df.copy()
    
    # Replace all forms of NA/null values with empty strings
    export_df = export_df.replace({
        pd.NA: "",
        "<NA>": "",
        "nan": "",
        "None": "",
        "NaT": ""
    })
    export_df = export_df.fillna("")
    
    # Convert all string columns to ensure no <NA> values remain
    for col in export_df.select_dtypes(include=['string', 'object']).columns:
        export_df[col] = export_df[col].astype(str).replace({"<NA>": "", "nan": "", "None": ""})
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Cleaned_Aligned")
    
    return output.getvalue()