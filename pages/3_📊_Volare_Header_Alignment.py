import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO
from openpyxl.styles import numbers

st.title("📊 VOLARE Header Alignment")
st.divider()

# --- Base Path ---
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
VOLARE_REFERENCE_PATH = os.path.join("utils", "reference", "volare_bpi_cards_xdays-header.xlsx")

# --- Helper: Normalize Column Names ---
def normalize(col: str):
    return str(col).strip().lower().replace("_", " ").replace("-", " ")

# --- Load VOLARE Header Mapping ---
def load_volare_header_mapping():
    ref_df = pd.read_excel(VOLARE_REFERENCE_PATH, header=None)
    header_mapping = {}

    # Each column = 1 header + alternate names
    for col in ref_df.columns:
        col_values = ref_df[col].dropna().astype(str).str.strip().tolist()
        if not col_values:
            continue
        main_header = col_values[0]
        alternatives = col_values[1:]
        header_mapping[main_header] = alternatives

    return header_mapping

# --- Helper: Map Headers Dynamically ---
def map_headers(df: pd.DataFrame, header_mapping: dict):
    mapped_cols = {}
    normalized_cols = {normalize(c): c for c in df.columns}

    for main_header, alt_names in header_mapping.items():
        all_aliases = [normalize(a) for a in alt_names + [main_header]]
        for alias in all_aliases:
            if alias in normalized_cols:
                mapped_cols[main_header] = normalized_cols[alias]
                break

    return mapped_cols

# Load mapping
VOLARE_HEADER_MAPPING = load_volare_header_mapping()
VOLARE_HEADERS = list(VOLARE_HEADER_MAPPING.keys())

# Define columns that should preserve leading zeros
TEXT_COLUMNS = [
    "CUST_ID", "CUSTOMER_ID", "CUSTOMER ID", "CUST ID",
    "CONTACT NUMBER", "CONTACT_NUMBER", "MOBILE", "PHONE",
    "MOBILE_NO", "CELL_PHONE", "TEL_NO", "TELEPHONE"
]

# Define date-related keywords
DATE_KEYWORDS = [
    "date", "dob", "as of", "due", "birth", "expiry", "exp", 
    "cust opn", "open", "d_cust_opn"
]

# File upload for alignment
uploaded_file = st.file_uploader(
    "📤 Upload Excel file for VOLARE header alignment",
    type=["xlsx", "xls"]
)

if uploaded_file:
    # Read the uploaded file with text columns as strings
    try:
        # First read to detect columns
        temp_df = pd.read_excel(uploaded_file)
        
        # Identify text columns
        dtype_dict = {}
        for col in temp_df.columns:
            col_normalized = normalize(col)
            if any(normalize(text_col) in col_normalized for text_col in TEXT_COLUMNS):
                dtype_dict[col] = str
        
        # Re-read with proper dtypes
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, dtype=dtype_dict)
        
    except Exception as e:
        st.error(f"Error reading file: {e}")
        df = pd.DataFrame()
    
    if not df.empty:
        st.success("✅ File successfully loaded!")
        
        # Show original data preview
        st.subheader("🧾 Original Data Preview")
        st.dataframe(df.head(10), use_container_width=True)
        
        st.divider()
        
        # Header Alignment section
        st.subheader("📑 Align to VOLARE Header")
        st.caption("Automatically align headers using reference file mapping.")

        if st.button("➡️ Proceed to Header Alignment"):
            header_map = map_headers(df, VOLARE_HEADER_MAPPING)
            missing_headers = [h for h in VOLARE_HEADERS if h not in header_map]

            if missing_headers:
                st.warning("⚠️ Missing columns were added as blank: " + ", ".join(missing_headers))

            aligned_df = pd.DataFrame()
            for h in VOLARE_HEADERS:
                if h in header_map:
                    aligned_df[h] = df[header_map[h]]
                else:
                    aligned_df[h] = ""  # Add blank column for missing

            # === Preserve leading zeros for text columns ===
            for col in aligned_df.columns:
                col_normalized = normalize(col)
                if any(normalize(text_col) in col_normalized for text_col in TEXT_COLUMNS):
                    # Convert to string and preserve leading zeros
                    aligned_df[col] = aligned_df[col].astype(str).str.strip()
                    # Clean up invalid values
                    aligned_df[col] = aligned_df[col].replace(
                        ["nan", "None", "NaT", ""], ""
                    )
                    # Ensure it stays as string
                    aligned_df[col] = aligned_df[col].apply(
                        lambda x: x if pd.isna(x) or x == "" or str(x).strip() == "" else str(x)
                    )

            # === Special handling for CUST_ID (add leading zero if missing) ===
            if "CUST_ID" in aligned_df.columns:
                def format_cust_id(value):
                    if pd.isna(value) or str(value).strip() == "" or str(value) == "nan":
                        return ""
                    val_str = str(value).strip()
                    # If it's numeric and doesn't start with 0, add leading 0
                    if val_str.replace('.', '').replace('-', '').isdigit() and not val_str.startswith('0'):
                        # Remove decimal if present
                        if '.' in val_str:
                            val_str = str(int(float(val_str)))
                        return f"0{val_str}"
                    return val_str
                
                aligned_df["CUST_ID"] = aligned_df["CUST_ID"].apply(format_cust_id)

            # Clean MEMO_LINE
            if "MEMO_LINE" in aligned_df.columns:
                aligned_df["MEMO_LINE"] = aligned_df["MEMO_LINE"].astype(str).str.replace(r"[^A-Za-z0-9 ]", "", regex=True)
                aligned_df["MEMO_LINE"] = aligned_df["MEMO_LINE"].replace("nan", "")

            # === Convert date columns to MM/DD/YYYY format (as TEXT) ===
            def format_date_text(value):
                if pd.isna(value) or str(value).strip() == "" or str(value) == "nan":
                    return ""
                try:
                    dt = pd.to_datetime(value, errors="coerce")
                    if pd.isna(dt):
                        return str(value)  # leave non-date as-is
                    return dt.strftime("%m/%d/%Y")
                except Exception:
                    return str(value)

            # Identify and format date columns
            date_columns = []
            for col in aligned_df.columns:
                if any(word in col.lower() for word in DATE_KEYWORDS):
                    aligned_df[col] = aligned_df[col].apply(format_date_text)
                    date_columns.append(col)

            aligned_df = aligned_df.fillna("")

            st.success("✅ Headers successfully aligned!")
            st.dataframe(aligned_df.head(10), use_container_width=True)

            # === Prepare download with proper Excel formatting ===
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                aligned_df.to_excel(writer, index=False, sheet_name="VOLARE_Aligned")
                
                # Get the worksheet to format columns
                worksheet = writer.sheets["VOLARE_Aligned"]
                
                # Format text columns (preserve leading zeros)
                for idx, col in enumerate(aligned_df.columns, start=1):
                    col_normalized = normalize(col)
                    
                    # Check if it's a text column (ID, phone, etc.)
                    if any(normalize(text_col) in col_normalized for text_col in TEXT_COLUMNS):
                        # Set entire column to text format
                        for row in range(2, len(aligned_df) + 2):
                            cell = worksheet.cell(row=row, column=idx)
                            cell.number_format = numbers.FORMAT_TEXT
                    
                    # Check if it's a date column
                    elif any(word in col.lower() for word in DATE_KEYWORDS):
                        # Set date columns as TEXT format (mm/dd/yyyy already formatted as string)
                        for row in range(2, len(aligned_df) + 2):
                            cell = worksheet.cell(row=row, column=idx)
                            cell.number_format = numbers.FORMAT_TEXT
                            # Alternative: use date format if you want Excel to recognize as date
                            # cell.number_format = 'MM/DD/YYYY'
            
            output.seek(0)
            
            st.download_button(
                label="📤 Download VOLARE Aligned File",
                data=output.getvalue(),
                file_name="VOLARE_aligned_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.info(f"ℹ️ Date columns formatted as text (mm/dd/yyyy): {', '.join(date_columns) if date_columns else 'None detected'}")
            
    else:
        st.warning("⚠️ The uploaded file appears to be empty.")
else:
    st.info("👆 Please upload an Excel file to align with VOLARE headers.")