import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from utils.helpers import read_excel_file
import re

st.title("📱 TU Numbers Splitter")
st.divider()

uploaded_file = st.file_uploader(
    "📤 Upload an Excel file (single file)",
    type=["xlsx", "xls"]
)

if not uploaded_file:
    st.info("👆 Please upload an Excel file to begin.")
else:
    df, _ = read_excel_file(uploaded_file)
    if df is None:
        st.error("❌ Failed to read uploaded file.")
    else:
        if df.empty:
            st.warning("⚠️ The uploaded file appears to be empty.")
        else:
            st.success("✅ File loaded successfully")

            cols = list(df.columns)

            col1, col2 = st.columns([1, 1])
            with col1:
                ch_col = st.selectbox("Select CH CODE column", cols, index=0)
            with col2:
                # choose a sensible default if 'TU NUMBERS' exists
                default_idx = 1
                if "TU NUMBERS" in cols:
                    default_idx = cols.index("TU NUMBERS")
                elif "TU_NUMBER" in cols:
                    default_idx = cols.index("TU_NUMBER")
                tu_col = st.selectbox("Select TU NUMBERS column", cols, index=default_idx)

            separator = st.text_input("Separator (single character or string)", value=";")
            options_col1, options_col2 = st.columns(2)
            with options_col1:
                trim_spaces = st.checkbox("Trim spaces from each contact", value=True)
            with options_col2:
                remove_empty = st.checkbox("Remove empty contacts", value=True)

            st.write("---")

            if st.button("➡️ Split TU NUMBERS"):
                working = df[[ch_col, tu_col]].copy()

                # Normalize column names in the output
                working.columns = ["CH CODE", "TU_RAW"]
                

                # Helper to split and clean
                def split_and_clean(cell):
                    """
                    Split and clean TU numbers in a single cell.
                    Rules:
                    - Split using the chosen separator.
                    - Remove any part containing full emails.
                    - Extract numeric sequences (6+ digits) if part of a text.
                    - Normalize Philippine numbers:
                        * +63 / 63 / 0063 → 0xxxxxxxxxx
                        * Add leading 0 if only 10 digits.
                    - Remove duplicates and empty entries.
                    """
                    if pd.isna(cell):
                        return []

                    s = str(cell)
                    parts = s.split(separator)
                    cleaned = []
                    seen = set()  # for avoiding duplicates

                    def normalize_number(num_str):
                        """Normalize and return a cleaned mobile number."""
                        digits = re.sub(r"\D", "", num_str)

                        if not digits:
                            return None

                        # --- Keep landline numbers (7 to 9 digits) ---
                        if 7 <= len(digits) <= 9 and not digits.startswith(("0", "63")):
                            return digits

                        # Handle international formats +63, 0063, or 63
                        if digits.startswith("63") and len(digits) >= 12:
                            digits = "0" + digits[-10:]
                        elif digits.startswith("0") and len(digits) == 11:
                            pass  # already valid
                        elif len(digits) == 10 and not digits.startswith("0"):
                            digits = "0" + digits

                        # Return only if it looks valid
                          # --- Final validation ---
                        if len(digits) == 11 and digits.startswith("0"):
                            return digits
                        elif 7 <= len(digits) <= 9:  # still allow landlines here as fallback
                            return digits
                        else:
                            return None

                    for p in parts:
                        if trim_spaces:
                            p = p.strip()
                        if remove_empty and (p == "" or str(p).lower() in ["nan", "none", "nat"]):
                            continue

                        # Remove or skip emails
                        if re.search(r"\S+@\S+\.[A-Za-z]{2,}", p):
                            p = re.sub(r"\S+@\S+\.[A-Za-z]{2,}", "", p).strip()
                            if not p:
                                continue

                        # Extract first long numeric sequence (e.g., TU 12272024-rejected)
                        match = re.search(r"(\d{6,})", p)
                        if match:
                            num_candidate = match.group(1)
                            normalized = normalize_number(num_candidate)
                            if normalized and normalized not in seen:
                                cleaned.append(normalized)
                                seen.add(normalized)
                            continue

                        # Otherwise try to normalize directly
                        normalized = normalize_number(p)
                        if normalized and normalized not in seen:
                            cleaned.append(normalized)
                            seen.add(normalized)

                    return cleaned


                # Apply splitting
                working["_splits"] = working["TU_RAW"].apply(split_and_clean)

                # Determine maximum number of contacts
                max_contacts = working["_splits"].apply(len).max()
                if pd.isna(max_contacts) or max_contacts is None:
                    max_contacts = 0
                max_contacts = int(max_contacts)

                # Build output rows
                out_rows = []
                for _, row in working.iterrows():
                    ch = row["CH CODE"]
                    splits = row["_splits"] if isinstance(row["_splits"], list) else []
                    # pad with empty strings
                    padded = splits + [""] * (max_contacts - len(splits))
                    r = {"CH CODE": ch}
                    for i, val in enumerate(padded, start=1):
                        r[f"CONTACT {i}"] = val
                    out_rows.append(r)

                if not out_rows:
                    st.warning("⚠️ No TU numbers found to split.")
                else:
                    out_df = pd.DataFrame(out_rows)

                    # Show preview
                    st.subheader("🧾 Split TU Numbers Preview")
                    st.dataframe(out_df.head(20), width="stretch")

                    # Prepare excel download
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        out_df.to_excel(writer, index=False, sheet_name="TU_Numbers_Split")
                    b = output.getvalue()

                    st.download_button(
                        label="📥 Download TU Numbers File",
                        data=b,
                        file_name="tu_numbers_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.success(f"✅ Done — produced {len(out_df)} rows with up to {max_contacts} CONTACT columns.")

            else:
                # show a small preview of the columns selected
                st.subheader("🧾 Sample of selected columns")
                st.dataframe(df[[ch_col, tu_col]].head(10), width="stretch")