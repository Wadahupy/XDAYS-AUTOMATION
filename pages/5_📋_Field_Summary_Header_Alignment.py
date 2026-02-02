import streamlit as st
import pandas as pd
from io import BytesIO
import re
from datetime import datetime

#--- Page Setup ---
st.set_page_config(
    page_title="Field Summary Header Alignment",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📋 Field Summary Header Alignment")
st.caption("Upload an Excel file to automatically align headers based on predefined mapping rules.")
st.divider()

#--- Header Mapping Rules ---
HEADER_MAPPING = {
    "BANK": "BANK",
    "PLACEMENT": "PLACEMENT",
    "CH CODE": "CH CODE",
    "ACCOUNT NUMBER": "ACCOUNT NUMBER",
    "CH NAME": "CH NAME",
    "ADD TYPE": "ADD TYPE",
    "ADDRESS": "ADDRESS",
    "DL TYPE": "DL TYPE",
    "OB/PRINCIPAL": "OB/PRINCIPAL",
    "AREA": "AREA",
    "MUNICIPALITY": "MUNICIPALITY",
    "BRGY": "BRGY",
    "FINAL AREA": "FINAL AREA",
    "DL TYPE 2": "DL TYPE",
    "REFERENCE CODE": "REFERENCE CODE",
    "AUTOFIELD DATE": "DL DATE",
    "OB": "OB/PRINCIPAL",
    "ENDO DATE": "ENDO DATE",
    "EMAIL SUBJECT": "EMAIL SUBJECT",
    "REMARKS": "REMARKS",
    "CONCAT CYCLE": "CONCAT CYCLE",
    "CYCLE": "CYCLE",
}

# Main output headers (in order)
MAIN_HEADERS = [
    "BANK",
    "PLACEMENT",
    "CH CODE",
    "ACCOUNT NUMBER",
    "CH NAME",
    "ADD TYPE",
    "ADDRESS",
    "DL TYPE",
    "OB/PRINCIPAL",
    "AREA",
    "MUNICIPALITY",
    "BRGY",
    "FINAL AREA",
    "DL TYPE (2)",
    "REFERENCE CODE",
    "AUTOFIELD DATE",
    "OB",
    "ENDO DATE",
    "EMAIL SUBJECT",
    "REMARKS",
    "CONCAT CYCLE",
    "CYCLE",
]

# --- Helper Functions ---
def normalize_header(header: str) -> str:
    """Normalize header: strip spaces, uppercase, remove dots."""
    normalized = str(header).strip().upper()
    normalized = normalized.replace(".", "")
    return normalized


def format_date_column(df: pd.DataFrame, col_name: str) -> pd.DataFrame:
    """Convert date columns to mm/dd/yyyy if possible."""
    if col_name not in df.columns:
        return df

    try:
        if df[col_name].isna().all():
            return df

        temp = pd.to_datetime(df[col_name], errors='coerce')
        if temp.notna().sum() > 0:
            df[col_name] = temp.dt.strftime('%m/%d/%Y')
    except:
        pass

    return df


def find_mapped_header(uploaded_header: str, mapping_rules: dict) -> str:
    """Return mapped header or None."""
    normalized_uploaded = normalize_header(uploaded_header)

    for key, value in mapping_rules.items():
        if normalize_header(key) == normalized_uploaded:
            return value

    return None


#--- Main Alignment Function ---
def align_headers(df: pd.DataFrame, mapping_rules: dict, main_headers: list) -> tuple:
    """
    Align headers to main output headers.
    - Rename matched columns to main headers
    - Add missing main headers as blank columns
    - Output in exact order of main_headers
    """
    uploaded_headers = df.columns.tolist()
    alignment_map = {}
    matched_count = 0
    unmatched_headers = []

    # Determine mappings for uploaded headers
    for header in uploaded_headers:
        mapped = find_mapped_header(header, mapping_rules)
        if mapped:
            alignment_map[header] = mapped
            matched_count += 1
        else:
            alignment_map[header] = None
            unmatched_headers.append(header)

    # Rename matched headers
    rename_dict = {old: new for old, new in alignment_map.items() if new is not None}
    aligned_df = df.rename(columns=rename_dict)

    # Remove unmapped columns
    aligned_df = aligned_df.drop(columns=unmatched_headers)

    # Add missing main headers as blank columns
    for main_header in main_headers:
        if main_header not in aligned_df.columns:
            aligned_df[main_header] = ""

    # Reorder to match MAIN_HEADERS order
    aligned_df = aligned_df[main_headers]

    # Find which main headers are missing
    present_headers = set([v for v in alignment_map.values() if v is not None])
    missing_headers = [h for h in main_headers if h not in present_headers]

    # Format date columns after renaming
    date_columns = ["AUTOFIELD DATE", "ENDO DATE"]
    for col in date_columns:
        aligned_df = format_date_column(aligned_df, col)

    report = {
        "uploaded_headers": uploaded_headers,
        "alignment_map": alignment_map,
        "matched_count": matched_count,
        "unmatched_headers": unmatched_headers,
        "missing_headers": missing_headers,
        "extra_headers": unmatched_headers,
    }

    return aligned_df, report


def generate_excel_download(df: pd.DataFrame) -> bytes:
    """Convert DataFrame to Excel bytes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Aligned Data")
    output.seek(0)
    return output.getvalue()

#--- File Upload ---
uploaded_file = st.file_uploader("📁 Upload Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")

        st.success(f"✅ File uploaded successfully! ({len(df)} rows, {len(df.columns)} columns)")
        st.divider()

        # ALIGNMENT PROCESSING
        aligned_df, report = align_headers(df, HEADER_MAPPING, MAIN_HEADERS)

      # --- Preview Alignment Results ---
        st.subheader("🔄 Header Alignment Preview")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### Detected Headers (Uploaded)")

            alignment_data = []
            for uploaded_header in report["uploaded_headers"]:
                mapped = report["alignment_map"].get(uploaded_header)
                status = "✅" if mapped else "❌"
                alignment_data.append({
                    "Uploaded Header": uploaded_header,
                    "Status": status,
                    "Mapped To": mapped or "[UNMAPPED]"
                })

            alignment_df = pd.DataFrame(alignment_data)

            def highlight_status(val):
                if val == "✅":
                    return "background-color: #90EE90; color: black;"
                elif val == "❌":
                    return "background-color: #FFB6C1; color: black;"
                return ""

            styled_alignment = alignment_df.style.applymap(
                highlight_status, subset=["Status"]
            )
            st.dataframe(styled_alignment, use_container_width=True, hide_index=True)

        with col2:
            st.markdown("### Alignment Summary")
            st.metric("Matched Headers", report["matched_count"])
            st.metric("Unmapped Headers", len(report["unmatched_headers"]))
            st.metric("Extra Columns", len(report["extra_headers"]))

        st.divider()

        # --- Validation Report ---
        st.subheader("⚠️ Validation Report")

        v1, v2 = st.columns(2)

        with v1:
            st.warning("**Unmapped Headers:**")
            for h in report["unmatched_headers"]:
                st.write(f"• {h}")

        with v2:
            st.info("**Extra Columns (left as-is):**")
            for h in report["extra_headers"]:
                st.write(f"• {h}")

        st.divider()

        # --- Aligned Data Preview ---
        st.subheader("📊 Aligned Data Preview")
        st.dataframe(aligned_df, use_container_width=True)

        st.divider()

        # DOWNLOAD
        excel_bytes = generate_excel_download(aligned_df)
        st.download_button(
            label="⬇️ Download Aligned Excel File",
            data=excel_bytes,
            file_name=f"aligned_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")

else:
    st.info("👈 Upload an Excel file to begin.")
