import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Field Autostat",
    layout="wide",
    initial_sidebar_state="expanded"
)


st.title("🛵 FIELD RESULT Automation Export")
st.caption("Upload an create excel file for FIELD RESULT Autostat.")
st.divider()

# ---- Upload File ----
uploaded_file = st.file_uploader("Upload your CSV/Excel file", type=["csv", "xlsx"])

if uploaded_file:
    # Read file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file, sheet_name="RESULT")

    # ---- Required Columns Check ----
    required_cols = [
        'PLACEMENT', 'DATE', 'TIME', 'status', 'Message',
        'chcode', 'PTP-Date', 'PTP AMOUNT'
    ]

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}")
        st.stop()

    # ---- Date Handling ----
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    df['TIME'] = pd.to_timedelta(df['TIME'].astype(str), errors='coerce')
    df = df.dropna(subset=['DATE'])

    st.success("File uploaded successfully!")

    # ---- Filters ----
    placement_options = df['PLACEMENT'].dropna().unique()
    min_date = df['DATE'].dt.date.min()
    max_date = df['DATE'].dt.date.max()

    with st.form("filter_form"):
        selected_placement = st.multiselect(
            "Select PLACEMENT(s):",
            options=placement_options
        )

        selected_date_range = st.date_input(
            "Select Date Range",
            value=[min_date, max_date],
            min_value=min_date,
            max_value=max_date
        )

        submit = st.form_submit_button("Run Automation")

    # ---- Apply Filters & Build Final Output ----
    if submit:
        if isinstance(selected_date_range, (list, tuple)) and len(selected_date_range) == 2:
            start_date, end_date = selected_date_range
        else:
            st.error("Please select a valid start and end date.")
            st.stop()

        filtered_df = df[
            (df['PLACEMENT'].isin(selected_placement)) &
            (df['DATE'].dt.date >= start_date) &
            (df['DATE'].dt.date <= end_date)
        ].copy()

        # ---- Action Status Mapping ----
        action_status_map = {
            'PTP': 'PTP - PASTDUE_FLD VST',
            'NEG': 'FIELD VISIT - VST RESULT_NEG',
            'POS': 'FIELD VISIT - VST RESULT_POS CH',
            'POS W/C': 'FIELD VISIT - VST RESULT_POS CH',
            'TP': 'FIELD VISIT - VST RESULT_POS THIRD PARTY'
        }

        # ---- Build Output DataFrame ----
        output = pd.DataFrame()

        output['chcode'] = filtered_df['chcode']
        output['Action Status'] = filtered_df['status'].map(action_status_map)

        # Remark Date = DATE + TIME
        remark_datetime = filtered_df['DATE'] + filtered_df['TIME']
        output['Remark Date'] = remark_datetime.dt.strftime("%m/%d/%Y %I:%M:%S %p")

        # PTP Date
        output['PTP Date'] = pd.to_datetime(
            filtered_df['PTP-Date'], errors='coerce'
        ).dt.strftime("%m/%d/%Y")

        # Reason For Default (blank)
        output['Reason For Default'] = ""

        # Field Visit Date
        output['Field Visit Date'] = remark_datetime.dt.strftime("%m/%d/%Y")

        # Remark
        output['Remark'] = (
            "| Remarks: " +
            output['Field Visit Date'] +
            " FIELD RESULT : " +
            filtered_df['Message'].astype(str)
        )

        # Next Call Date (blank)
        output['Next Call Date'] = ""

        # PTP Amount
        output['PTP Amount'] = filtered_df['PTP AMOUNT'].replace([0, ""], pd.NA)

        # Claim Paid Amount (blank)
        output['Claim Paid Amount'] = ""

        # Remark By (Placement based)
        output['Remark By'] = filtered_df['PLACEMENT'].map({
            'BPI CARDS 30 DPD': 'RALOPE',
            'BPI CARDS XDAYS': 'MMMEJIA'
        })

        # Remaining blank columns
        output['Phone No.'] = ""
        output['Relation'] = ""
        output['Claim Paid Date'] = ""

        # ---- Display Result ----
        st.subheader("Automation Output")
        st.dataframe(output, use_container_width=True)

        # ---- Download Excel ----
        def to_excel(df):
            output_buffer = BytesIO()
            with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="UPLOAD")
            return output_buffer.getvalue()

        excel_data = to_excel(output)

        st.download_button(
            "Download Automation File",
            data=excel_data,
            file_name="FIELD_RESULT_AUTOMATION.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.caption(f"Rows generated: **{len(output):,}**")