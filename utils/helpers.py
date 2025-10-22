import pandas as pd
import re
import io
import streamlit as st
from typing import Tuple, List, Dict, Set
import numpy as np
import msoffcrypto
import os
from io import BytesIO
from datetime import datetime

# Configure pandas
pd.set_option('future.no_silent_downcasting', True)

# ----------------------------
# 📁 PATHS
# ----------------------------
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
REFERENCE_PATH = os.path.join(BASE_PATH, "reference", "bpi_cards_xdays-header.xlsx")


# ----------------------------
# 📘 LOAD REFERENCE HEADER FILE
# ----------------------------
def load_reference_headers() -> Tuple[List[str], Dict[str, str]]:
    """
    Load main headers and alternates from reference/bpi_cards_xdays-header.xlsx
    Expected structure:
        Each column = 1 header group
        Row 1 = main header
        Row 2+ = alternate names
    """
    if not os.path.exists(REFERENCE_PATH):
        st.error(f"❌ Reference header file not found at: {REFERENCE_PATH}")
        st.stop()

    try:
        ref_df = pd.read_excel(REFERENCE_PATH, header=0, engine="openpyxl")

        header_mapping = {}
        for col in ref_df.columns:
            alternates = ref_df[col].dropna().tolist()
            if alternates:
                main_header = alternates[0]
                alt_names = alternates[1:]
                for alt in alt_names:
                    header_mapping[str(alt).strip().upper()] = main_header

        # Keep column sequence (main headers in order)
        column_sequence = [ref_df[col].iloc[0] for col in ref_df.columns if pd.notna(ref_df[col].iloc[0])]
        return column_sequence, header_mapping

    except Exception as e:
        st.error(f"❌ Error loading header reference: {e}")
        st.stop()


# ----------------------------
# 📄 READ EXCEL FILE
# ----------------------------
def read_excel_file(uploaded_file) -> Tuple[pd.DataFrame, List[str]]:
    """Read Excel file (handles encryption if necessary)."""
    try:
        file_name = getattr(uploaded_file, "name", "")
        file_ext = file_name.split(".")[-1].lower()
        decrypted = BytesIO()

        try:
            office_file = msoffcrypto.OfficeFile(uploaded_file)
            if office_file.is_encrypted():
                st.warning(f"🔒 File '{file_name}' is password-protected.")
                password = st.text_input(f"Enter password for {file_name}:", type="password", key=file_name)
                if not password:
                    st.stop()
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
                st.success(f"✅ Successfully decrypted {file_name}.")
            else:
                uploaded_file.seek(0)
                decrypted = BytesIO(uploaded_file.read())
        except Exception:
            uploaded_file.seek(0)
            decrypted = BytesIO(uploaded_file.read())

        engine = "xlrd" if file_ext == "xls" else "openpyxl"
        df = pd.read_excel(decrypted, engine=engine)
        df.columns = [str(c).strip() for c in df.columns]

        # Standardize strings
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).replace(['nan', 'None', 'NaT', 'null'], np.nan)

        return df, list(df.columns)

    except Exception as e:
        st.error(f"❌ Error reading file {file_name}: {str(e)}")
        return None, []


# ----------------------------
# 🧹 CLEAN DATA (UPDATED FULL VERSION)
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
    if cycle.isdigit() and len(cycle) == 1:
        cycle = f"0{cycle}"
    df["COLLECTION_CYCLE"] = cycle

    # =====================
    # 🧹 Clean CUST_ID
    # =====================
    if "CUST_ID" in df.columns:
        df["CUST_ID"] = df["CUST_ID"].astype(str).str.strip()
        df["CUST_ID"] = df["CUST_ID"].replace(["nan", "None", "NaT", "", "??", "????"], np.nan)

        df["TEMP_CUST_ID"] = df["CUST_ID"].apply(lambda x: f"0{x}" if pd.notna(x) and not str(x).startswith("0") else x)

        df["TEMP_CUST_ID"] = df["TEMP_CUST_ID"].apply(
            lambda x: str(int(float(x))) if isinstance(x, str) and re.match(r"^\d+\.0+$", x) else x
        )

        df["CUST_ID"] = df["TEMP_CUST_ID"]
        df.drop(columns=["TEMP_CUST_ID"], inplace=True)

        # =====================
    # 📱 Clean MOBILE_NO (detect and normalize any format)
    # =====================
    if "MOBILE_NO" in df.columns:
        df["MOBILE_NO"] = df["MOBILE_NO"].astype(str).str.strip()

        def normalize_mobile(num):
            if pd.isna(num):
                return np.nan

            # Remove all non-digit characters
            num = re.sub(r"\D", "", str(num))

            # Remove leading country codes and redundant prefixes
            # Examples it handles:
            # +63917..., 63917..., 00963917..., 0917..., 917...
            if num.startswith("00"):
                num = num[2:]
            if num.startswith("63"):
                num = "0" + num[2:]
            elif num.startswith("9") and len(num) == 10:
                num = "0" + num

            # Only keep if it’s a valid PH mobile number
            if re.fullmatch(r"09\d{9}", num):
                return num

            return np.nan

        df["MOBILE_NO"] = df["MOBILE_NO"].apply(normalize_mobile)


        # =====================
        # 📅 Clean & standardize date columns
        # =====================
        DATE_COLUMNS = [
            "LAST_PAYMENT_DATE", "LAST_CONTACT_DATE", "PTP DATE", "PTP_DATE",
            "BIRTHDATE", "LAST DUE DATE", "LAST_DUE_DATE",
            "TPAP DD", "TPAP_DD", "D_CUST_OPN", "D CUST OPEN"
        ]

        def normalize_name(name: str) -> str:
            """Normalize column names for flexible matching."""
            return str(name).strip().lower().replace("_", " ")

        def convert_any_date(value):
            """Convert Excel serials or date strings to mm/dd/yyyy."""
            if pd.isna(value):
                return np.nan
            try:
                # Handle Excel serial dates (within reasonable range)
                if isinstance(value, (int, float)) and 1 < value < 60000:
                    value = pd.to_datetime("1899-12-30") + pd.to_timedelta(value, "D")
                else:
                    value = pd.to_datetime(value, errors="coerce")
                return value.strftime("%m/%d/%Y") if pd.notna(value) else np.nan
            except Exception:
                return np.nan

        for col in df.columns:
            if normalize_name(col) in [normalize_name(dc) for dc in DATE_COLUMNS]:
                df[col] = df[col].apply(convert_any_date)






    # =====================
    # 🔤 String columns
    # =====================
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
            invalid_values = ["nan", "None", "NaT", "????", "??"]
            if col in ["OFC", "HOME", "OFFICE_PH", "HOME_PH"]:
                invalid_values.extend(["0", "00"])

            df[col] = df[col].astype(str)
            df[col] = df[col].apply(lambda x: np.nan if str(x).strip() in invalid_values else str(x).strip())
            if pd.notna(df[col]).any():
                df[col] = df[col].astype('string')

    return df


# ----------------------------
# ⚙️ ALIGN HEADERS (using reference)
# ----------------------------
def align_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Align headers using reference/header_alignment.xlsx."""
    column_sequence, header_mapping = load_reference_headers()

    def normalize(h): return re.sub(r'[\s_]+', '', str(h).strip().upper())
    normalized_mapping = {normalize(k): v for k, v in header_mapping.items()}

    new_cols = []
    for col in df.columns:
        norm_col = normalize(col)
        mapped_col = normalized_mapping.get(norm_col, col.strip())
        new_cols.append(mapped_col)

    df.columns = new_cols

    # Add missing columns from master sequence
    for col in column_sequence:
        if col not in df.columns:
            df[col] = np.nan

    return df.reindex(columns=column_sequence)


# ----------------------------
# 💾 DOWNLOAD EXCEL
# ----------------------------
def generate_excel_download(df: pd.DataFrame, filename: str) -> bytes:
    """Generate downloadable Excel."""
    export_df = df.replace({
        pd.NA: "", "<NA>": "", "nan": "", "None": "", "NaT": ""
    }).fillna("")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Aligned")
    return output.getvalue()
