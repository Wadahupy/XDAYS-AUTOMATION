# pages/2_📂_File_Consolidation.py
import streamlit as st
import pandas as pd
import os
from io import BytesIO
from datetime import datetime
import xlwt
from utils.helpers import read_excel_file

# === Page setup ===
st.title("📂 Excel File Consolidation")
st.caption("Merge multiple Excel files into a single consolidated file without altering any data or formatting.")
st.divider()

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

    TEXT_COLUMNS = [
        "CUSTOMER_ID", "CUSTOMER ID", "CUST_ID", "CUST ID",
        "CONTACT NUMBER", "CONTACT_NUMBER", "MOBILE", "PHONE",
        "MOBILE_ALS", "MOBILE_ALFES", "PRIMARY_NO_ALS", 
        "BUS_NO_ALS", "LANDLINE_NO_ALS", "LAN"
    ]

    for file in uploaded_files:
        try:
            temp_df = pd.read_excel(file)

            dtype_dict = {}
            for col in temp_df.columns:
                if any(text_col.upper() in col.upper() for text_col in TEXT_COLUMNS):
                    dtype_dict[col] = str

            file.seek(0)
            df = pd.read_excel(file, dtype=dtype_dict)

        except Exception as e:
            st.warning(f"⚠️ Error reading {file.name}: {e}")
            stats["skipped_files"].append(file.name)
            continue

        if df is None or df.empty:
            stats["skipped_files"].append(file.name)
            continue

        df.columns = [str(c).strip() for c in df.columns]

        # Ensure text columns are strings and preserve leading zeros
        for col in df.columns:
            if any(text_col.upper() in col.upper() for text_col in TEXT_COLUMNS):
                df[col] = df[col].astype(str).str.strip().replace("nan", "")

        all_dfs.append(df)
        stats["processed_files"] += 1
        stats["total_rows"] += len(df)

    if not all_dfs:
        return None, stats

    consolidated = pd.concat(all_dfs, ignore_index=True, sort=False)

    for col in consolidated.columns:
        if any(text_col.upper() in col.upper() for text_col in TEXT_COLUMNS):
            consolidated[col] = consolidated[col].astype(str).str.strip().replace("nan", "")

    return consolidated, stats

# --- When Files Uploaded ---
if uploaded_files:
    consolidated_df, stats = simple_consolidate_files(uploaded_files)

    if consolidated_df is not None and not consolidated_df.empty:
        st.success("✅ Files successfully consolidated!")

        # --- Summary ---
        st.subheader("📊 Consolidation Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Files", stats["total_files"])
        col2.metric("Total Rows", stats["total_rows"])
        col3.metric("Total Columns", len(consolidated_df.columns))

        if stats["skipped_files"]:
            st.warning(f"⚠️ Skipped files: {', '.join(stats['skipped_files'])}")

        st.divider()
        st.subheader("🧾 Consolidated Data Preview")
        st.dataframe(consolidated_df.head(10), use_container_width=True)

        # === XLS GENERATION ===
        total_accounts = len(consolidated_df)
        today = datetime.now().strftime("%m%d%Y")
        output_filename = f"bpi_xdays_consolidated_{total_accounts}_{today}.xls"

        output = BytesIO()
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Consolidated")

        TEXT_COLUMNS = [
            "CUSTOMER_ID", "CUSTOMER ID", "CUST_ID", "CUST ID",
            "CONTACT NUMBER", "CONTACT_NUMBER", "MOBILE", "PHONE",
            "MOBILE_ALS", "MOBILE_ALFES", "PRIMARY_NO_ALS",
            "BUS_NO_ALS", "LANDLINE_NO_ALS", "LAN"
        ]

        text_style = xlwt.XFStyle()
        text_style.num_format_str = "@"

        # Write headers
        for col_index, col_name in enumerate(consolidated_df.columns):
            sheet.write(0, col_index, col_name)

        # Write data rows
        for row_index, row in consolidated_df.iterrows():
            for col_index, col_name in enumerate(consolidated_df.columns):
                value = row[col_name]

                # Force TRUE empty cells
                if pd.isna(value) or value in ["nan", "NaN", None]:
                    value = ""

                # Apply text-only formatting
                if any(t.upper() in col_name.upper() for t in TEXT_COLUMNS):
                    sheet.write(row_index + 1, col_index, str(value), text_style)
                else:
                    sheet.write(row_index + 1, col_index, value)

        # Save XLS to buffer
        workbook.save(output)
        output.seek(0)

        # === Download button ===
        st.download_button(
            label="📥 Download Consolidated XLS File",
            data=output.getvalue(),
            file_name=output_filename,
            mime="application/vnd.ms-excel"
        )

        st.divider()

    else:
        st.warning("⚠️ No valid data found for consolidation.")

else:
    st.info("👆 Please upload Excel files to consolidate.")
