# pages/1_📑_Header_Alignment.py
import streamlit as st
from utils.helpers import read_excel_file, align_headers, generate_excel_download, clean_data
from datetime import datetime
import re

st.title("📑 Excel Header Auto Alignment")
st.divider()

uploaded_file = st.file_uploader("Upload an Excel file (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file:
    df, headers = read_excel_file(uploaded_file)

    if df is not None:
        # 🧹 Clean data automatically
        cleaned_df = clean_data(df, uploaded_file.name)

        st.subheader("🧽 Cleaned Data Preview")
        st.dataframe(cleaned_df.head(5), width="stretch")

        st.divider()
        st.subheader("⚙️ Auto Alignment")

        aligned_df = align_headers(cleaned_df)

        st.success("✅ Columns automatically aligned to the correct sequence.")
        st.dataframe(aligned_df.head(5), width="stretch")

        st.divider()
        st.subheader("🧩 Column Selection")

        # 🧠 Let user choose which columns to include
        all_columns = list(aligned_df.columns)
        selected_columns = st.multiselect(
            "Select which columns to include in the final output:",
            options=all_columns,
            default=all_columns
        )

        # 🔍 Identify excluded columns
        excluded_columns = [col for col in all_columns if col not in selected_columns]

        # 📋 Show excluded columns (if any)
        st.markdown("#### ❌ Columns Not Included")
        if excluded_columns:
            st.write(", ".join(excluded_columns))
        else:
            st.info("✅ All columns are currently included in the output.")

        st.divider()

        # ✅ Filter the aligned DataFrame based on user's selection
        if selected_columns:
            final_df = aligned_df[selected_columns]
            st.success(f"✅ {len(selected_columns)} columns selected for export.")
            st.dataframe(final_df.head(5), width="stretch")

            # --------------------------
            # 📊 DASHBOARD SUMMARY
            # --------------------------
            st.divider()
            st.subheader("📈 Data Overview Dashboard")

            col1, col2 = st.columns(2)

            with col1:
                st.metric("Total Rows", f"{len(final_df):,}")
            with col2:
                st.metric("Total Columns", f"{len(final_df.columns):,}")

            # Optional breakdown: show per-column missing %
            with st.expander("🔍 View Column Completeness Details"):
                missing_info = final_df.isnull().sum().reset_index()
                missing_info.columns = ["Column", "Missing Count"]
                missing_info["Missing %"] = (
                    missing_info["Missing Count"] / len(final_df) * 100
                ).round(2)
                st.dataframe(missing_info, width="stretch")

            # Extract cycle number from uploaded filename
            match = re.search(r"c(\d+)", uploaded_file.name.lower())
            cycle = match.group(1).upper() if match else "N/A"

            # Get today's date formatted as MM-DD-YYYY
            today_str = datetime.now().strftime("%m-%d-%Y")

            # Create final export filename
            export_filename = f"BPI_XDAYS_C{cycle}_HEADER_ALIGNED_{today_str}.xlsx"

            # Download button with formatted name
            st.download_button(
                label=f"📥 Download {export_filename}",
                data=generate_excel_download(final_df, export_filename),
                file_name=export_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Please select at least one column to include in the output.")

else:
    st.info("👆 Upload an Excel file to begin.")
