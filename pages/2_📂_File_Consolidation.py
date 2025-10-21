# pages/2_📂_File_Consolidation.py
import streamlit as st
import pandas as pd
import numpy as np
import os
from io import BytesIO
from utils.helpers import generate_excel_download, read_excel_file

st.title("📂 Excel File Consolidation")
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

# Load mapping
VOLARE_HEADER_MAPPING = load_volare_header_mapping()
VOLARE_HEADERS = list(VOLARE_HEADER_MAPPING.keys())

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

# --- File Upload ---
uploaded_files = st.file_uploader(
    "📤 Upload multiple Excel files for consolidation",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# --- Consolidation Logic ---
def simple_consolidate_files(uploaded_files):
    if not uploaded_files:
        return None, {}

    stats = {"total_files": len(uploaded_files), "processed_files": 0, "skipped_files": [], "total_rows": 0}
    all_dfs = []

    for file in uploaded_files:
        df, _ = read_excel_file(file)
        if df is None or df.empty:
            stats["skipped_files"].append(file.name)
            continue

        df.columns = [str(c).strip() for c in df.columns]
        all_dfs.append(df)
        stats["processed_files"] += 1
        stats["total_rows"] += len(df)

    if not all_dfs:
        return None, stats

    consolidated = pd.concat(all_dfs, ignore_index=True, sort=False).fillna("")
    return consolidated, stats

# --- When Files Uploaded ---
if uploaded_files:
    consolidated_df, stats = simple_consolidate_files(uploaded_files)

    if consolidated_df is not None:
        st.success("✅ Files successfully consolidated!")

        # Summary
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

        # Download Consolidated
        st.download_button(
            label="📥 Download Consolidated File",
            data=generate_excel_download(consolidated_df, "consolidated_output.xlsx"),
            file_name="consolidated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # --- Header Alignment ---
        st.subheader("📑 Align to VOLARE Header")
        st.caption("Automatically align headers using reference file mapping.")

        if st.button("➡️ Proceed to Header Alignment"):
            header_map = map_headers(consolidated_df, VOLARE_HEADER_MAPPING)
            missing_headers = [h for h in VOLARE_HEADERS if h not in header_map]

            if missing_headers:
                st.warning("⚠️ Missing columns were added as blank: " + ", ".join(missing_headers))

            aligned_df = pd.DataFrame()
            for h in VOLARE_HEADERS:
                if h in header_map:
                    aligned_df[h] = consolidated_df[header_map[h]]
                else:
                    aligned_df[h] = ""  # Add blank column for missing

            # === Ensure CUST_ID has leading zero and is text ===
            if "CUST_ID" in aligned_df.columns:
                aligned_df["CUST_ID"] = aligned_df["CUST_ID"].astype(str).str.strip()
                aligned_df["CUST_ID"] = aligned_df["CUST_ID"].replace(
                    ["nan", "None", "NaT", "", "??", "????"], np.nan
                )
                aligned_df["CUST_ID"] = aligned_df["CUST_ID"].apply(
                    lambda x: f"0{str(int(float(x))) if str(x).replace('.', '').isdigit() else str(x)}"
                    if pd.notna(x) and not str(x).startswith("0")
                    else str(x)
                )
                aligned_df["CUST_ID"] = aligned_df["CUST_ID"].fillna("").astype(str)
                # --- Clean MEMO_LINE like Excel TEXTJOIN formula ---
            if "MEMO_LINE" in aligned_df.columns:
                aligned_df["MEMO_LINE"] = aligned_df["MEMO_LINE"].astype(str).str.replace(r"[^A-Za-z0-9 ]", "", regex=True)

            
            # 3️⃣ Convert all date-like columns to text format MM/DD/YYYY
            def format_date_text(value):
                if pd.isna(value) or str(value).strip() == "":
                    return ""
                try:
                    # Try parsing any date-like format
                    dt = pd.to_datetime(value, errors="coerce")
                    if pd.isna(dt):
                        return str(value)  # leave non-date as-is
                    return dt.strftime("%m/%d/%Y")
                except Exception:
                    return str(value)

            # Automatically detect and format columns that look like dates
            for col in aligned_df.columns:
                if any(word in col.lower() for word in ["date", "dob", "as of", "due"]):
                    aligned_df[col] = aligned_df[col].apply(format_date_text)



            aligned_df = aligned_df.fillna("")

            st.success("✅ Headers successfully aligned (with blanks for missing columns)!")
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
    st.info("👆 Please upload Excel files to consolidate.")
