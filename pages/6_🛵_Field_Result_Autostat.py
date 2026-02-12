import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.title("FIELD RESULT Automation Export")

uploaded_file = st.file_uploader("Upload your CSV/Excel file", type=["csv", "xlsx"])

if uploaded_file:
    USE_COLS = [
        'PLACEMENT', 'DATE', 'TIME', 'status', 'Message',
        'chcode', 'PTP-Date', 'PTP AMOUNT'
    ]

    # ---- Read File (FASTER) ----
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, usecols=USE_COLS)
    else:
        df = pd.read_excel(uploaded_file, sheet_name="RESULT", usecols=USE_COLS)

    # ---- Datetime Handling (ONE PASS) ----
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    df['TIME'] = pd.to_timedelta(df['TIME'].astype(str), errors='coerce')
    df['REMARK_DATETIME'] = df['DATE'] + df['TIME']

    df = df.dropna(subset=['DATE'])

    st.success("File uploaded successfully!")

    # ---- Filters ----
    placement_options = df['PLACEMENT'].dropna().unique()
    min_date = df['DATE'].dt.date.min()
    max_date = df['DATE'].dt.date.max()

    with st.form("filter_form"):
        selected_placement = st.multiselect("Select PLACEMENT(s):", placement_options)
        selected_date_range = st.date_input(
            "Select Date Range",
            [min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )
        submit = st.form_submit_button("Run Automation")

    if submit:
        start_date, end_date = selected_date_range

        # ---- FAST filtering ----
        mask = (
            df['PLACEMENT'].isin(selected_placement) &
            df['DATE'].dt.date.between(start_date, end_date)
        )
        f = df.loc[mask]

        # ---- Mappings ----
        action_status_map = {
            'PTP': 'PTP - PASTDUE_FLD VST',
            'NEG': 'FIELD VISIT - VST RESULT_NEG',
            'POS': 'FIELD VISIT - VST RESULT_POS CH',
            'POS W/C': 'FIELD VISIT - VST RESULT_POS CH',
            'TP': 'FIELD VISIT - VST RESULT_POS THIRD PARTY'
        }

        remark_by_map = {
            'BPI CARDS 30 DPD': 'RALOPE',
            'BPI CARDS XDAYS': 'MMMEJIA',
            'BPI PL XDAYS': 'JGCELIZ',
            'BPI PL 30DPD': 'APLADIP',
            'BPI PL 60DPD': 'NNLUQUIAS',
            'BPI RBANK CARDS 30DPD': 'MIMFERNANDEZ',
            'BPI RBANK PL SL': 'MIMFERNANDEZ',
            'CARDS RECOV': 'JDCULDORA',
            'RECOVERY': "JDCULDORA",
            'BUCKET 4': "JDCULDORA"
        }

        # ---- Build Output (FAST) ----
        output = pd.DataFrame({
            'chcode': f['chcode'],
            'Action Status': f['status'].map(action_status_map),
            'Remark Date': f['REMARK_DATETIME'].dt.strftime("%m/%d/%Y %I:%M:%S %p"),
            'PTP Date': pd.to_datetime(f['PTP-Date'], errors='coerce').dt.strftime("%m/%d/%Y"),
            'Reason For Default': "",
            'Field Visit Date': f['REMARK_DATETIME'].dt.strftime("%m/%d/%Y"),
            'Remark': "| Remarks: " +
                      #f['REMARK_DATETIME'].dt.strftime("%m/%d/%Y"# +
                      " FIELD RESULT : " + f['Message'].astype(str),
            'Next Call Date': "",
            'PTP Amount': f['PTP AMOUNT'].replace([0, ""], pd.NA),
            'Claim Paid Amount': "",
            'Remark By': f['PLACEMENT'].map(remark_by_map),
            'Phone No.': "",
            'Relation': "",
            'Claim Paid Date': ""
        })

        st.subheader("Automation Output")
        st.dataframe(output, use_container_width=True)

        # ---- Cached Excel Writer ----
        @st.cache_data
        def to_excel_fast(df):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="UPLOAD")
            return buffer.getvalue()

        excel_data = to_excel_fast(output)

        filename = (
            f"FIELD_RESULT_{datetime.now():%Y-%m-%d}_"
            f"{'_'.join(selected_placement) if selected_placement else 'ALL'}.xlsx"
        )

        st.download_button(
            "Download Automation File",
            excel_data,
            filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.caption(f"Rows generated: **{len(output):,}**")