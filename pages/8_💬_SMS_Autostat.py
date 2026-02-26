import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta

st.title("CYN BLASTING Automation (Active File Updated Headers)")
st.caption("Upload an CYN Blasting file to automatically generate a remarks.")
st.divider()

# ===== FILE UPLOADS =====
blast_file = st.file_uploader("Upload CYN Blasting File", type=["csv", "xlsx"])
active_file = st.file_uploader("Upload Active File", type=["csv", "xlsx"])

if blast_file and active_file:

    # ===== READ BLAST FILE =====
    if blast_file.name.endswith(".csv"):
        df = pd.read_csv(blast_file)
    else:
        df = pd.read_excel(blast_file)

    # ===== READ ACTIVE FILE =====
    if active_file.name.endswith(".csv"):
        active_df = pd.read_csv(active_file)
    else:
        active_df = pd.read_excel(active_file)

    st.success("Both files uploaded successfully!")

    # ===== REQUIRED COLUMN CHECK =====
    required_blast = ['Loan Number', 'MOBILE_NO', 'TEMPLATE']
    required_active = ['CUST_ID', 'CH CODE']

    missing_blast = [c for c in required_blast if c not in df.columns]
    missing_active = [c for c in required_active if c not in active_df.columns]

    if missing_blast:
        st.error(f"Blast file missing columns: {missing_blast}")
        st.stop()

    if missing_active:
        st.error(f"Active file missing columns: {missing_active}")
        st.stop()

    # ===== USER INPUT =====
    with st.form("cyn_form"):
        base_date = st.date_input("Base Date (C3)", datetime.today())
        hour_offset = st.number_input("Hour Offset (E3)", value=0)
        submit = st.form_submit_button("Run Automation")

    if submit:

        # ===== CLEAN MATCH COLUMNS (VERY IMPORTANT) =====
        df['Loan Number'] = df['Loan Number'].astype(str).str.strip()
        active_df['CUST_ID'] = active_df['CUST_ID'].astype(str).str.strip()

        # ===== REMARK DATETIME =====
        remark_datetime = datetime.combine(base_date, datetime.min.time()) + timedelta(hours=hour_offset)

        # ===== LOOKUP (MATCH Loan Number → CUST_ID) =====
        merged = df.merge(
            active_df[['CUST_ID', 'CH CODE']],
            how='left',
            left_on='Loan Number',
            right_on='CUST_ID'
        )

        # ===== BUILD OUTPUT =====
        output = pd.DataFrame()

        # ACCOUNT NUMBER from CH CODE
        output['ACCOUNT NUMBER'] = merged['CH CODE']

        # CHECKER
        output['CHECKER'] = output['ACCOUNT NUMBER'].apply(
            lambda x: "for checking" if pd.isna(x) else "ok to proceed"
        )

        output['Action Status'] = "SMS SENT - BLAST SMS"
        output['Remark Date'] = remark_datetime.strftime("%m/%d/%Y %I:%M:%S %p")
        output['PTP Date'] = ""
        output['Reason For Default'] = ""
        output['Field Visit Date'] = ""#remark_datetime.strftime("%m/%d/%Y")

        output['Phone No.'] = "0" + df['MOBILE_NO'].astype(str).str[-10:]

        output['Remark'] = (
            output['Phone No.'] +
            " - SMS Sent SMS Blast: " +
            df['TEMPLATE'].astype(str) +
            " template"
        )

        output['Next Call Date'] = ""
        output['PTP Amount'] = ""
        output['Claim Paid Amount'] = ""
        output['Remark By'] = "SPMADRID"
        output['Relation'] = ""
        output['Claim Paid Date'] = ""

        # Filter OK TO PROCEED
        filtered = output[output['CHECKER'].str.lower() == "ok to proceed"]

        # Remove CHECKER before export
        final_output = filtered.drop(columns=['CHECKER'])

        # Ensure exact column order (including Phone No.)
        final_output = final_output[
            [
                'ACCOUNT NUMBER',
                'Action Status',
                'Remark Date',
                'PTP Date',
                'Reason For Default',
                'Field Visit Date',
                'Remark',
                'Next Call Date',
                'PTP Amount',
                'Claim Paid Amount',
                'Remark By',
                'Phone No.',
                'Relation',
                'Claim Paid Date'
            ]
        ]


        st.subheader("Automation Output Preview")
        st.dataframe(final_output, use_container_width=True)

        # ===== EXPORT =====
        @st.cache_data
        def to_excel(df):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="UPLOAD")
            return buffer.getvalue()

        excel_data = to_excel(final_output)

        filename = f"CYN_BLASTING_IMPORT_{datetime.now():%Y-%m-%d}.xlsx"

        st.download_button(
            "Download CYN BLASTING File",
            excel_data,
            filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ===== SUMMARY =====
        total = len(output)
        matched = len(final_output)
        unmatched = total - matched

        st.success(f"""
        ✅ Total Records: {total:,}
        ✅ Matched (OK TO PROCEED): {matched:,}
        ⚠️ For Checking: {unmatched:,}
        """)