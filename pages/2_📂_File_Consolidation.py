# pages/2_📂_File_Consolidation.py
import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO
from utils.helpers import generate_excel_download, read_excel_file

# === Page setup ===
st.title("📂 Excel File Consolidation")
st.caption("Merge multiple Excel files into a single consolidated file without altering any data or formatting.")
st.divider()

# --- Base Path ---
BASE_PATH = os.path.dirname(os.path.abspath(__file__))

# --- File Upload ---
uploaded_files = st.file_uploader(
    "📤 Upload multiple Excel files to consolidate",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# --- Consolidation Logic ---
def simple_consolidate_files(uploaded_files):
    if not uploaded_files:
        return None, {}

    stats = {
        "total_files": len(uploaded_files),
        "processed_files": 0,
        "skipped_files": [],
        "total_rows": 0
    }
    all_dfs = []

    # Columns that should preserve leading zeros (customize as needed)
    TEXT_COLUMNS = [
        "CUSTOMER_ID", "CUSTOMER ID", "CUST_ID", "CUST ID",
        "CONTACT NUMBER", "CONTACT_NUMBER", "MOBILE", "PHONE",
        "MOBILE_ALS", "MOBILE_ALFES", "PRIMARY_NO_ALS", 
        "BUS_NO_ALS", "LANDLINE_NO_ALS", "LAN"
    ]

    for file in uploaded_files:
        # Read with dtype=str for text columns to preserve leading zeros
        try:
            # First, read to detect all columns
            temp_df = pd.read_excel(file)
            
            # Identify which text columns exist in this file
            dtype_dict = {}
            for col in temp_df.columns:
                col_upper = str(col).strip().upper()
                # Check if column name matches any text column pattern
                if any(text_col.upper() in col_upper for text_col in TEXT_COLUMNS):
                    dtype_dict[col] = str
            
            # Re-read with proper dtypes
            file.seek(0)  # Reset file pointer
            df = pd.read_excel(file, dtype=dtype_dict)
            
        except Exception as e:
            st.warning(f"⚠️ Error reading {file.name}: {e}")
            stats["skipped_files"].append(file.name)
            continue
        
        if df is None or df.empty:
            stats["skipped_files"].append(file.name)
            continue

        # Keep headers exactly as they appear
        df.columns = [str(c).strip() for c in df.columns]
        
        # Ensure text columns are strings and preserve leading zeros
        for col in df.columns:
            col_upper = col.upper()
            if any(text_col.upper() in col_upper for text_col in TEXT_COLUMNS):
                # Convert to string and preserve leading zeros
                df[col] = df[col].astype(str).str.strip()
                # Replace 'nan' with empty string
                df[col] = df[col].replace('nan', '')
        
        all_dfs.append(df)
        stats["processed_files"] += 1
        stats["total_rows"] += len(df)

    if not all_dfs:
        return None, stats

    # Combine all data as-is (no cleaning or formatting)
    consolidated = pd.concat(all_dfs, ignore_index=True, sort=False)
    
    # Final pass: ensure all text columns are strings in consolidated file
    for col in consolidated.columns:
        col_upper = col.upper()
        if any(text_col.upper() in col_upper for text_col in TEXT_COLUMNS):
            consolidated[col] = consolidated[col].astype(str).str.strip()
            consolidated[col] = consolidated[col].replace('nan', '')
    
    return consolidated, stats

# --- When Files Uploaded ---
if uploaded_files:
    consolidated_df, stats = simple_consolidate_files(uploaded_files)

    if consolidated_df is not None and not consolidated_df.empty:
        st.success("✅ Files successfully consolidated!")

        # --- Summary ---
        st.subheader("📊 Consolidation Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Files", stats.get("total_files", 0))
        col2.metric("Total Rows", stats.get("total_rows", 0))
        col3.metric("Total Columns", len(consolidated_df.columns))

        if stats.get("skipped_files"):
            st.warning(f"⚠️ Skipped files: {', '.join(stats['skipped_files'])}")

        st.divider()
        st.subheader("🧾 Consolidated Data Preview")
        st.dataframe(consolidated_df.head(10), use_container_width=True)

        # --- Download Button with proper Excel formatting ---
        # Create Excel file with text format for specific columns
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            consolidated_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            # Get the worksheet
            worksheet = writer.sheets['Consolidated']
            
            # Format specific columns as text to preserve leading zeros
            from openpyxl.styles import numbers
            
            TEXT_COLUMNS = [
                "CUSTOMER_ID", "CUSTOMER ID", "CUST_ID", "CUST ID",
                "CONTACT NUMBER", "CONTACT_NUMBER", "MOBILE", "PHONE",
                "MOBILE_ALS", "MOBILE_ALFES", "PRIMARY_NO_ALS", 
                "BUS_NO_ALS", "LANDLINE_NO_ALS", "LAN"
            ]
            
            # Find column indices for text columns
            for idx, col in enumerate(consolidated_df.columns, start=1):
                col_upper = col.upper()
                if any(text_col.upper() in col_upper for text_col in TEXT_COLUMNS):
                    # Set column to text format
                    for row in range(2, len(consolidated_df) + 2):  # +2 because row 1 is header
                        cell = worksheet.cell(row=row, column=idx)
                        cell.number_format = numbers.FORMAT_TEXT
        
        output.seek(0)
        
        st.download_button(
            label="📥 Download Consolidated File",
            data=output.getvalue(),
            file_name="consolidated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

    else:
        st.warning("⚠️ No valid data found for consolidation.")
else:
    st.info("👆 Please upload Excel files to consolidate.")