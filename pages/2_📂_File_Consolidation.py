# pages/2_📂_File_Consolidation.py
import streamlit as st
import pandas as pd
from utils.helpers import generate_excel_download, read_excel_file
import numpy as np

st.title("📂 Excel File Consolidation")
st.divider()

uploaded_files = st.file_uploader(
    "Upload multiple Excel files for consolidation",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# === Official VOLARE upload header sequence ===
VOLARE_HEADERS = [
    "CH CODE", "CUST_ID", "CUST_NAME", "AOD", "QUEUE", "OFC", "HOME", "MOBILE_NO", "TU",
    "ADDRESS", "EMAIL", "GENDER", "BOS", "MAD", "XDAYS", "LAST PAYMENT AMOUNT",
    "LAST_PAYMENT_DATE", "DELIQUENCY_STRING", "ADA_ACCOUNT",
    "DEBIT_AMOUNT_PREFERENCE", "UNIT_CODE", "HO FLAG", "BIRTHDAY",
    "BLOCK_CODE", "MEMO_LINE", "D_CUST_OPN", "LAST_DUE_DATE",
    "SPOUSE NUMBER", "TYPE"
]

# === Alternative / Alias headers ===
HEADER_ALIASES = {
    "CH CODE": ["ch code"],
    "CUST_ID": ["cust_id", "customer id"],
    "CUST_NAME": ["cust_name", "customer name"],
    "AOD": ["aod"],
    "QUEUE": ["queue", "level/cycle"],
    "OFC": ["ofc"],
    "HOME": ["home_ph", "contact", "phone number"],
    "MOBILE_NO": ["mobile_no", "mobile number", "MOBILE_NU"],
    "TU": ["tu", "TU_NUMBERS"],
    "ADDRESS": ["address", "PRIMARY ADDRESS", "primary address"],
    "EMAIL": ["email"], 
    "GENDER": ["gender"],
    "BOS": ["bos"],
    "MAD": ["mad"],
    "XDAYS": ["xdays"],
    "LAST PAYMENT AMOUNT": ["last payment amount", "LAST_PAYMENT"],
    "LAST_PAYMENT_DATE": ["last_payment_date", "LAST_PAYMENT_DATE"],
    "DELIQUENCY_STRING": ["delinquency_string", "DELINQUENC", "delinquenc"],
    "ADA_ACCOUNT": ["ada_account"],
    "DEBIT_AMOUNT_PREFERENCE": ["debit_amount_preference"],
    "UNIT_CODE": ["unit_code", "FINNONE", "finnone"],
    "HO FLAG": ["ho flag"],
    "BIRTHDAY": ["birthdate", "birthday"],
    "BLOCK_CODE": ["block_code"],
    "MEMO_LINE": ["memo_line"],
    "D_CUST_OPN": ["d_cust_opn"],
    "LAST_DUE_DATE": ["last due date"],
    "SPOUSE NUMBER": ["spouse number"],
    "TYPE": ["type"]
}

def normalize(col: str):
    """Simplify column names for comparison"""
    return str(col).strip().lower().replace("_", " ")

def simple_consolidate_files(uploaded_files):
    """
    Consolidate files WITHOUT applying header alignment or column filtering.
    Simply reads all files and concatenates them, preserving ALL columns.
    """
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
        df, _ = read_excel_file(file)
        if df is None or df.empty:
            stats["skipped_files"].append(file.name)
            continue
        
        # Basic cleanup only - no header alignment
        # Clean up any whitespace in column names
        df.columns = [str(c).strip() for c in df.columns]
        
        all_dfs.append(df)
        stats["processed_files"] += 1
        stats["total_rows"] += len(df)
    
    if not all_dfs:
        st.error("❌ No valid files to consolidate.")
        return None, stats
    
    # Concatenate all dataframes - this will include ALL columns from all files
    consolidated = pd.concat(all_dfs, ignore_index=True, sort=False)
    
    # Fill NaN values in columns that didn't exist in some files
    consolidated = consolidated.fillna("")
    
    return consolidated, stats

def map_headers(df: pd.DataFrame):
    """Map alternative headers to official VOLARE headers"""
    mapped_cols = {}
    normalized_cols = {normalize(c): c for c in df.columns}

    for official, aliases in HEADER_ALIASES.items():
        # include official header itself as alias
        all_aliases = [normalize(a) for a in aliases + [official]]
        for alias in all_aliases:
            if alias in normalized_cols:
                mapped_cols[official] = normalized_cols[alias]
                break

    return mapped_cols

if uploaded_files:
    # Use the simple consolidation function instead
    consolidated_df, stats = simple_consolidate_files(uploaded_files)

    if consolidated_df is not None:
        st.success("✅ Files successfully consolidated!")

        # === Summary ===
        st.subheader("📊 Consolidation Summary")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Files Uploaded", stats.get('total_files', 0))
        with col2:
            st.metric("Total Rows Consolidated", stats.get('total_rows', 0))
        with col3:
            st.metric("Total Columns Detected", len(consolidated_df.columns))

        # Show skipped files if any
        if stats.get('skipped_files'):
            st.warning(f"⚠️ Skipped files: {', '.join(stats['skipped_files'])}")

        st.divider()
        st.subheader("🧾 Consolidated Data Preview")
        st.dataframe(consolidated_df.head(10), use_container_width=True)

        # === Download consolidated ===
        st.download_button(
            label="📥 Download Consolidated Excel File",
            data=generate_excel_download(consolidated_df, "consolidated_output.xlsx"),
            file_name="consolidated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # === Header Alignment Section ===
        st.subheader("📑 VOLARE Header Alignment")
        st.caption("Align the consolidated file to the official VOLARE Upload header format (case-insensitive + alias support).")

        if st.button("➡️ Proceed to Header Alignment"):
            header_map = map_headers(consolidated_df)

            missing_headers = [h for h in VOLARE_HEADERS if h not in header_map]
            if missing_headers:
                st.error("⚠️ Missing columns: " + ", ".join(missing_headers))
                st.warning("Please make sure all required headers exist in your uploaded data.")
            else:
                aligned_df = pd.DataFrame()
                for h in VOLARE_HEADERS:
                    aligned_df[h] = consolidated_df[header_map[h]]

                aligned_df = aligned_df.fillna("")

                st.success("✅ Headers successfully aligned to VOLARE format!")
                st.dataframe(aligned_df.head(10), use_container_width=True)

                st.download_button(
                    label="📤 Download VOLARE Aligned File",
                    data=generate_excel_download(aligned_df, "VOLARE_aligned_output.xlsx"),
                    file_name="VOLARE_aligned_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    else:
        st.warning("⚠️ No valid data found for consolidation.")
else:
    st.info("👆 Please upload multiple Excel files to consolidate.")