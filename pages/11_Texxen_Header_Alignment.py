import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📂 Header Automation Tool")

# --- Header Mapping Rules ---
HEADER_MAPPING = {
    "CUST_ID": "accountNumber",
    "LAN": "accountNumber",
    "CH CODE": "chCode",
    "CH_CODE": "chCode",
    "CUST_NAME": "name",
    "NAME": "name",
    "BIRTHDATE": "birthday",
    "AOD": "ob",
    "PAYOFF AMOUNT": "ob",
    "PRINCIPAL": "principal",
    "ENDO DATE": "endoDate",
    "DATE REFERRED": "endoDate",
    "LAST DUE DATE": "cutOff",
    "CTL4": "placement",
    "UNIT_CODE": "productType",
    "DPD": "dpd",
    "QUEUE": "level",
    "TCL": "creditLimit",
    "MAD": "installmentAmount",
    "LAST PAYMENT AMOUNT": "lastPaymentAmount",
    "LAST_PAYMENT_DATE": "lastPaymentDate",
    "EMAIL": "email",
    "EMAIL_ALS": "email",
    "ADDRESS": "address1",
    "MOBILE_NO": "phone1",
    "MOBILE_ALS": "phone1",
    "HOME": "phone2",
    "MOBILE_ALFES": "phone2",
    "OFC": "phone3",
    "PRIMARY_NO_ALS": "phone3",
    "BUS_NO_ALS": "phone4",
    "LANDLINE_NO_ALS": "phone5"
}

# --- CUSTOM_FIELDS per placement ---
CUSTOM_FIELDS_MAPPING = {
    "BPI CARDS XDAYS SL": [
        "GENDER", "DELINQUENCY_STRING", "ADA_ACCOUNT",
        "DEBIT_AMOUNT_PREFERENCE", "HO FLAG",
        "BLOCK_CODE", "MEMO_LINE", "D_CUST_OPN", "PDA", "assignedAgent", "assignedTeam"
    ],
    "BPI AUTO CURING SL": [
        "PAST DUE", "UNIT", "LPC", "ADA SHORTAGE", "assignedAgent",	"assignedTeam"
    ]
}

# --- Placement Selection ---
placement_name = st.selectbox(
    "Select Placement",
    options=list(CUSTOM_FIELDS_MAPPING.keys())
)

uploaded_file = st.file_uploader(
    "Upload Excel or CSV file",
    type=["xlsx", "xls", "csv"]
)



if uploaded_file:

    # --- Read file ---
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    # --- Normalize headers ---
    df.columns = df.columns.str.strip().str.upper()

    # --- Detect Truly Empty Columns (100% blank or NaN) ---
    empty_columns = []

    for col in df.columns:
        series = df[col]
        is_blank = series.isna() | (series.astype(str).str.strip() == "")
        if is_blank.all():
            empty_columns.append(col)

    # --- Map Non-Empty Standard Headers (KEEP DUPLICATES SAFELY) ---
    mapped_columns = []
    column_counts = {}

    for col in df.columns:
        if col in HEADER_MAPPING and col not in empty_columns:
            std_col = HEADER_MAPPING[col]

            # Handle duplicates safely
            if std_col in column_counts:
                column_counts[std_col] += 1
                new_col_name = f"{std_col}_{column_counts[std_col]}"
            else:
                column_counts[std_col] = 1
                new_col_name = std_col

            mapped_columns.append(
                pd.DataFrame({new_col_name: df[col]})
            )

    if not mapped_columns:
        st.error("❌ No valid headers detected. File not processed.")
        st.stop()

    df_mapped = pd.concat(mapped_columns, axis=1)

    # --- Determine Custom Fields for Placement ---
    custom_fields_for_placement = CUSTOM_FIELDS_MAPPING.get(
        placement_name, []
    )

    # --- Add Non-Empty Custom Fields ---
    for cf_name in custom_fields_for_placement:
        cf_name = cf_name.strip().upper()
        if cf_name in df.columns and cf_name not in empty_columns:
            df_mapped[cf_name] = df[cf_name]

    # --- Clean accountNumber and phone columns ---
    def clean_number(x):
        if pd.isna(x) or str(x).strip() == "":
            return ""
        try:
            return "0" + str(int(float(x)))
        except:
            return str(x).strip()

    for col in df_mapped.columns:
        if col.startswith("accountNumber") or col.startswith("phone"):
            df_mapped[col] = df_mapped[col].apply(clean_number)


        # --- Auto Assign assignedTeam Based on Placement ---
    if placement_name == "BPI AUTO CURING SL":
        df_mapped["assignedTeam"] = "BPI AUTO CURING SL TEAM 1"

    elif placement_name == "BPI CARDS XDAYS SL":
        df_mapped["assignedTeam"] = "BPI CARDS XDAYS SL TEAM 1"

    # --- Ensure assignedAgent column exists ---
    if "assignedAgent" not in df_mapped.columns:
        df_mapped["assignedAgent"] = ""

    # --- Force Final Column Order ---
    FINAL_COLUMN_ORDER = [
        "accountNumber", "accountNumber_2", "accountNumber_3",
        "cardNumber", "chCode", "name", "birthday", "ob", "newOb",
        "principal", "totalBalance", "endoDate", "cutOff",
        "placement", "productType", "cycle", "dpd", "level",
        "loanAmount", "creditLimit", "installmentAmount",
        "interest", "lastPaymentAmount", "lastPaymentDate",
        "writeOffAmount", "writeOffDate", "employer",
        "email", "address1", "address2", "address3",
        "address4", "address5",
        "phone1", "phone2", "phone3", "phone4", "phone5",

        # --- BPI CARDS ---
        "GENDER", "DELINQUENCY_STRING", "ADA_ACCOUNT",
        "DEBIT_AMOUNT_PREFERENCE", "HO FLAG",
        "BLOCK_CODE", "MEMO_LINE", "D_CUST_OPN", "PDA",

        # --- BPI AUTO CURING ---
        "PAST DUE", "UNIT", "LPC", "ADA SHORTAGE",

        "CUSTOM_FIELD_10", "assignedAgent", "assignedTeam"
    ]

    ordered_cols = [col for col in FINAL_COLUMN_ORDER if col in df_mapped.columns]
    remaining_cols = [col for col in df_mapped.columns if col not in ordered_cols]
    df_mapped = df_mapped[ordered_cols + remaining_cols]

    # --- Display results ---
    st.success("✅ Headers mapped successfully.")
    st.write("Returned Columns:")
    st.write(list(df_mapped.columns))
    st.dataframe(df_mapped.head())

    # --- Download File ---
    output_file = placement_name.replace(" ", "_") + "_TEXXEN.xlsx"

    buffer = BytesIO()
    df_mapped.to_excel(buffer, index=False)
    buffer.seek(0)

    st.download_button(
        label="Download Processed File",
        data=buffer,
        file_name=output_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )