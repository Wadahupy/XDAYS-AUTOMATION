# pages/1_📑_Header_Alignment.py
import streamlit as st
from utils.helpers import read_excel_file, align_headers, generate_excel_download

st.title("📑 Excel Header Auto Alignment")
st.divider()

uploaded_file = st.file_uploader("Upload an Excel file (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file:
    df, headers = read_excel_file(uploaded_file)

    if df is not None:
        st.subheader("🔍 Preview of Uploaded Data")
        st.dataframe(df.head(5), use_container_width=True)

        st.divider()
        st.subheader("⚙️ Auto Alignment")

        # 🔄 Automatically align headers to master column sequence
        aligned_df = align_headers(df)

        st.success("✅ Columns automatically aligned to the correct sequence.")
        st.dataframe(aligned_df.head(5), use_container_width=True)

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
            st.dataframe(final_df.head(5), use_container_width=True)

            st.download_button(
                label="📥 Download Filtered Excel File",
                data=generate_excel_download(final_df, "filtered_auto_aligned_output.xlsx"),
                file_name="filtered_auto_aligned_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Please select at least one column to include in the output.")

else:
    st.info("👆 Upload an Excel file to begin.")
