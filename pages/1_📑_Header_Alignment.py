import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime
from io import BytesIO
from utils.helpers import read_excel_file, clean_data, generate_excel_download

# ============================================================
# PAGE CONFIG
# ============================================================
st.title("📑 Excel Header Auto Alignment")
st.caption("Automatically align uploaded file headers using the reference template in `utils/reference/bpi_cards_xdays-header.xlsx`.")
st.divider()

# ============================================================
# LOAD REFERENCE HEADER FILE
# ============================================================
REFERENCE_PATH = os.path.join("utils", "reference", "bpi_cards_xdays-header.xlsx")

if not os.path.exists(REFERENCE_PATH):
    st.error("❌ Reference file not found: `utils/reference/bpi_cards_xdays-header.xlsx`")
    st.stop()

reference_df = pd.read_excel(REFERENCE_PATH, engine="openpyxl")
reference_headers = reference_df.columns.tolist()

# Create header mapping (main header = first row, alternates = rest)
header_mapping = {}
for col in reference_df.columns:
    alternates = reference_df[col].dropna().tolist()
    if alternates:
        main_header = alternates[0]
        alt_names = [a.strip() for a in alternates[1:] if str(a).strip()]
        header_mapping[main_header] = alt_names


# ============================================================
# FUNCTIONS
# ============================================================
def get_mapped_header(column_name: str, header_mapping: dict):
    """Find mapped header for a given uploaded column."""
    normalized_col = str(column_name).strip().lower()
    for main_header, alt_names in header_mapping.items():
        all_names = [main_header.lower()] + [a.lower() for a in alt_names]
        if normalized_col in all_names:
            return main_header
    return None


def align_headers_with_reference(uploaded_df: pd.DataFrame, reference_headers: list, header_mapping: dict):
    """Align headers of uploaded file to reference headers."""
    aligned_df = pd.DataFrame(columns=reference_headers)
    not_found_columns = []

    for column in uploaded_df.columns:
        mapped_header = get_mapped_header(column, header_mapping)
        if mapped_header:
            aligned_df[mapped_header] = uploaded_df[column]
        else:
            not_found_columns.append(column)

    aligned_df = aligned_df.fillna("")
    return aligned_df, not_found_columns


def export_to_excel(dataframe, sheet_name="Aligned Data"):
    """Export dataframe to an in-memory Excel file."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


# ============================================================
# FILE UPLOAD
# ============================================================
uploaded_file = st.file_uploader("📂 Upload an Excel file (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file:
    df, headers = read_excel_file(uploaded_file)

    if df is not None:
        # 🧽 Clean uploaded data
        cleaned_df = clean_data(df, uploaded_file.name)

        st.subheader("🧹 Cleaned Data Preview")
        st.dataframe(cleaned_df.head(5), width='stretch')

        st.divider()
        st.subheader("⚙️ Aligning Headers with Reference Template")

        aligned_df, not_found_columns = align_headers_with_reference(cleaned_df, reference_headers, header_mapping)

        st.success("✅ Columns successfully aligned using reference header template.")
        st.dataframe(aligned_df.head(5), width='stretch')

        # ============================================================
        # 🧩 COLUMN SELECTION
        # ============================================================
        st.divider()
        st.subheader("🧩 Column Selection")

        all_columns = list(aligned_df.columns)
        selected_columns = st.multiselect(
            "Select columns to include in the final output:",
            options=all_columns,
            default=all_columns
        )

        excluded_columns = [col for col in all_columns if col not in selected_columns]

        # 📋 Show excluded columns
        st.markdown("#### ❌ Columns Not Included")
        if excluded_columns:
            st.write(", ".join(excluded_columns))
        else:
            st.info("✅ All columns are currently included in the output.")

        # ============================================================
        # ✅ FINAL FILTERED DATAFRAME
        # ============================================================
        if selected_columns:
            final_df = aligned_df[selected_columns]

            st.divider()
            st.subheader("📈 Data Overview Dashboard")

            col1, col2 = st.columns(2)
            col1.metric("Total Rows", f"{len(final_df):,}")
            col2.metric("Total Columns", f"{len(final_df.columns):,}")

            with st.expander("🔍 View Column Completeness Details"):
                missing_info = final_df.isnull().sum().reset_index()
                missing_info.columns = ["Column", "Missing Count"]
                missing_info["Missing %"] = (missing_info["Missing Count"] / len(final_df) * 100).round(2)
                st.dataframe(missing_info, width='stretch')

                # ➕ Show columns from raw file not included in reference
                if not_found_columns:
                    st.markdown("### 🚫 Columns from Raw File Not in Reference Template")
                    not_found_df = pd.DataFrame(not_found_columns, columns=["Unmapped Columns"])
                    st.dataframe(not_found_df, width='stretch')
                

            # ============================================================
            # 📥 EXPORT FINAL ALIGNED FILE
            # ============================================================
            match = re.search(r"c(\d+)", uploaded_file.name.lower())
            cycle = match.group(1).zfill(2) if match else "N/A"
            today_str = datetime.now().strftime("%m-%d-%Y")
            export_filename = f"BPI_XDAYS_C{cycle}_HEADER_ALIGNED_{today_str}.xlsx"

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
