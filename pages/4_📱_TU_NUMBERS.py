import streamlit as st
import pandas as pd
import re
import numpy as np
from io import BytesIO
from utils.helpers import read_excel_file

st.title("📞 TU Numbers Splitter & Validator")
st.write("""
This tool:
1. Splits TU numbers using your chosen separator  
2. Auto-formats and validates PH phone numbers (mobile or landline)  
3. Returns **one row per CH CODE** with columns **CONTACT 1, CONTACT 2, ...**  
4. Lets you download valid and invalid numbers separately  
""")
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
            """Remove all non-digit characters except '*'."""
            val = to_str(val)
            return re.sub(r"[^\d*]", "", val)

        def auto_format_number(val):
            """Auto-format PH phone numbers including *0, *09, *63 patterns."""
            val = clean_value(val)
            if val == "":
                return ""

            if val.startswith("*"):
                val = val[1:]  # remove leading *

            # *63XXXXXXXXX or 63XXXXXXXXXX → mobile
            if val.startswith("63") and len(val) >= 12:
                return "0" + val[-10:]

            # 9XXXXXXXXX → mobile
            if re.fullmatch(r"9\d{9}", val):
                return "0" + val

            # Starts with 0 → already local
            if re.fullmatch(r"0\d{10}", val):  # mobile
                return val
            if re.fullmatch(r"0\d{9}", val):   # landline
                return val

            # 10 digits → assume landline
            if re.fullmatch(r"\d{10}", val):
                return val

            # 7–9 digits → assume landline, add 0
            if re.fullmatch(r"\d{7,9}", val):
                return "0" + val

            # 11 digits but not starting with 0 → fix
            if re.fullmatch(r"\d{11}", val):
                return "0" + val[1:]

            return val

        def is_valid_phone(val):
            """Check if number is valid mobile (11 digits) or landline (10 digits)."""
            val = to_str(val)
            pattern_mobile = re.compile(r"^0\d{10}$")
            pattern_landline = re.compile(r"^0\d{9}$")
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
                    valid_row[f"CONTACT {i}"] = num
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
