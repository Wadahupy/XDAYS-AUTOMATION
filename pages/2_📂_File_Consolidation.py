# pages/2_📂_File_Consolidation.py
import streamlit as st
import pandas as pd
from utils.helpers import consolidate_files, generate_excel_download, read_excel_file

st.title("📂 Excel File Consolidation")
st.divider()

uploaded_files = st.file_uploader(
    "Upload multiple Excel files for consolidation",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    consolidated_df, stats = consolidate_files(uploaded_files)

    if consolidated_df is not None:
        st.success("✅ Files successfully consolidated!")

        # Dashboard
        st.subheader("📊 Consolidation Summary")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Files Uploaded", stats.get('total_files', 0))
        with col2:
            st.metric("Total Rows Consolidated", stats.get('total_rows', 0))
        with col3:
            st.metric("Total Columns Aligned", stats.get('aligned_columns', 0))

        if stats.get('missing_columns'):
            st.warning("Missing columns detected: " + ", ".join(sorted(stats['missing_columns'])))

        st.divider()
        st.subheader("🧾 Consolidated Data Preview")
        st.dataframe(consolidated_df.head(10), use_container_width=True)


        # Download button
        st.download_button(
            label="📥 Download Consolidated Excel File",
            data=generate_excel_download(consolidated_df, "consolidated_output.xlsx"),
            file_name="consolidated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ No valid data found for consolidation.")
else:
    st.info("👆 Please upload multiple Excel files to consolidate.")
