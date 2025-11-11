import streamlit as st
import pandas as pd
import re
import numpy as np
from io import BytesIO
from utils.helpers import read_excel_file

st.title("📞 TU Numbers Splitter & Validator")

st.divider()

uploaded_file = st.file_uploader("📤 Upload Excel file", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("👆 Please upload an Excel file to begin.")
else:
    df, _ = read_excel_file(uploaded_file)
    if df is None or df.empty:
        st.error("❌ Failed to read or empty Excel file.")
    else:
        st.success("✅ File loaded successfully")

        cols = list(df.columns)
        col1, col2 = st.columns([1, 1])
        with col1:
            ch_col = st.selectbox("Select CH CODE column", cols, index=0)
        with col2:
            tu_col = st.selectbox("Select TU NUMBERS column", cols, index=1)

        separator = st.text_input("Separator (e.g., ',', ';', '|')", value=";")

        st.write("---")

         # ---------------- Helper functions ---------------- #
        def to_str(val):
            if pd.isna(val):
                return ""
            return str(val).strip()

        def clean_value(val):
            """Clean unwanted characters and remove emails or text."""
            val = to_str(val)
            # Remove email-like substrings
            val = re.sub(r"\S+@\S+\.[A-Za-z]{2,}", "", val)
            # Remove words like 'none', 'null', 'nan'
            if val.lower() in ["none", "nan", "null", "nat"]:
                return ""
            # Keep only digits and optional *
            val = re.sub(r"[^\d*]", "", val)
            return val

        def auto_format_number(val):
            """Auto-format PH phone numbers including *0, *09, *63 patterns, and reject invalid ones."""
            val = clean_value(val)
            if val == "":
                return ""

            # Remove leading asterisk if present
            if val.startswith("*"):
                val = val[1:]

            # Reject obvious invalid prefixes like 00, 000
            if re.match(r"^00+", val):
                return ""

            # Handle country code format (63XXXXXXXXXX)
            if val.startswith("63") and len(val) == 12:
                val = "0" + val[2:]

            # Handle 9XXXXXXXXX → 09XXXXXXXXX (mobile)
            elif re.fullmatch(r"9\d{9}", val):
                val = "0" + val

            # Handle landline (7–9 digits) → add 0
            elif re.fullmatch(r"\d{7,9}", val):
                val = "0" + val

            # Keep as is if already starts with 0 and correct length
            elif re.fullmatch(r"0\d{9,10}", val):
                pass
            else:
                return ""

            return val

        def is_valid_phone(val):
            """Check if number is valid mobile (11 digits) or landline (10 digits)."""
            val = to_str(val)
            if not val or re.match(r"^00+", val):  # reject 00 or 000 prefixes
                return False

            pattern_mobile = re.compile(r"^09\d{9}$")  # e.g., 09171234567
            pattern_landline = re.compile(r"^0\d{8,9}$")  # e.g., 0286871817

            return bool(pattern_mobile.match(val) or pattern_landline.match(val))
        # ---------------- Main process ---------------- #
        if st.button("🚀 Process TU NUMBERS"):
            working = df[[ch_col, tu_col]].copy()
            working.columns = ["CH CODE", "TU_RAW"]

            all_valid = []
            all_invalid = []

            for _, row in working.iterrows():
                ch_code = row["CH CODE"]
                numbers_raw = to_str(row["TU_RAW"])
                numbers = [n.strip() for n in numbers_raw.split(separator) if n.strip() != ""]

                valid_nums = []
                invalid_nums = []

                for num in numbers:
                    formatted = auto_format_number(num)
                    if is_valid_phone(formatted):
                        valid_nums.append(formatted)
                    else:
                        invalid_nums.append(num)

                valid_row = {"CH CODE": ch_code}
                for i, num in enumerate(valid_nums, start=1):
                    valid_row[f"TU {i}"] = num
                all_valid.append(valid_row)

                for bad in invalid_nums:
                    all_invalid.append({"CH CODE": ch_code, "Invalid Value": bad})

            valid_df = pd.DataFrame(all_valid).fillna("")
            invalid_df = pd.DataFrame(all_invalid).fillna("")

            st.success("✅ Processing complete!")

            st.write(f"**✅ Valid CH CODEs: {len(valid_df)} | ❌ Invalid entries: {len(invalid_df)}**")
            st.write("### ✅ Valid Numbers (One Row per CH CODE)")
            st.dataframe(valid_df, width="stretch")

            st.write("### ❌ Invalid Numbers")
            st.dataframe(invalid_df, width="stretch")

            # --- Downloads ---
            valid_output = BytesIO()
            valid_df.to_excel(valid_output, index=False)
            st.download_button(
                label="💾 Download VALID Numbers",
                data=valid_output.getvalue(),
                file_name="valid_tu_numbers.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if not invalid_df.empty:
                invalid_output = BytesIO()
                invalid_df.to_excel(invalid_output, index=False)
                st.download_button(
                    label="💾 Download INVALID Numbers",
                    data=invalid_output.getvalue(),
                    file_name="invalid_tu_numbers.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.subheader("🧾 Sample of selected columns")
            st.dataframe(df[[ch_col, tu_col]].head(10), width="stretch")
