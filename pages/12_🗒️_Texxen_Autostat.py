import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO  # <- needed for in-memory Excel file

# --- Streamlit App ---
st.set_page_config(page_title="File Header Alignment", layout="wide")
st.title("File Header Alignment Tool")

# Upload the file
uploaded_file = st.file_uploader("Upload your CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file:
    # Load the file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    
    st.subheader("Original Data")
    st.dataframe(df.head())

    # --- Column Mapping ---
    header_mapping = {
        "accountNumber": "Account No.",
        "groupStatus": "Call Status",
        "status": "Status2",  # will split later
        "substatus": "Substatus",
        "contactSource": "Contact Source",
        "dispositionSource": "Disposition Source",
        "newAddress": "New Address",
        "newEmail": "New Email",
        "paymentAmount": "Claim Paid Amount",
        "ptpAmount": "PTP Amount",
        "phoneNumber": "Dialed Number",
        "rfd": "RFD Status",
        "startDate": "PTP Date",
        "endDate": "End Date",
        "barcodeDate": "Barcode Date",
        "agent": "Remark By",
        "notes": "Remark",
        "Product Type": "Product Type"
    }

    # Rename columns that exist in the uploaded file
    df = df.rename(columns={k: v for k, v in header_mapping.items() if k in df.columns})

    # --- Split status into Status2 and Substatus ---
    if "Status2" in df.columns:
        df[["Status2", "Substatus"]] = df["Status2"].astype(str).str.split("-", n=1, expand=True)
    
    # --- Data Cleaning Functions ---
    def clean_number(x):
        if pd.isna(x) or str(x).strip() == "":
            return ""
        try:
            return "0" + str(int(float(x)))
        except:
            return str(x).strip()   

    
    def format_date(x):
        """Format date to mm/dd/yyyy"""
        if pd.isna(x) or str(x).strip() == "":
            return ""
        try:
            # Try to parse the date
            if isinstance(x, str):
                # Try common date formats
                for fmt in ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]:
                    try:
                        parsed_date = pd.to_datetime(x, format=fmt)
                        return parsed_date.strftime("%m/%d/%Y")
                    except:
                        continue
                # If no format matched, try pandas default parsing
                parsed_date = pd.to_datetime(x)
            else:
                parsed_date = pd.to_datetime(x)
            return parsed_date.strftime("%m/%d/%Y")
        except:
            return str(x).strip()
    
    # Clean Account Number
    if "Account No." in df.columns:
        df["Account No."] = df["Account No."].apply(clean_number)
    
    # Clean Dialed Number
    if "Dialed Number" in df.columns:
        df["Dialed Number"] = df["Dialed Number"].apply(clean_number)
    
    # Format Date columns
    date_columns = ["PTP Date", "End Date", "Barcode Date"]
    for col in date_columns:
        if col in df.columns and col != "Barcode Date":  # Don't format Barcode Date as it's already set
            df[col] = df[col].apply(format_date)

    # --- Arrange columns based on MAIN HEADER order ---
    main_header = [
        "Account No.",
        "Call Status",
        "Status2",
        "Substatus",
        "Contact Source",
        "Disposition Source",
        "New Address",
        "New Email",
        "Claim Paid Amount",
        "PTP Amount",
        "Dialed Number",
        "RFD Status",
        "PTP Date",
        "End Date",
        "Barcode Date",
        "Remark By",
        "Remark",
        "Product Type"
    ]

    # --- Barcode Date Selection ---
    st.subheader("Set Barcode Date & Time")

    col1, col2 = st.columns(2)

    with col1:
        barcode_date = st.date_input(
            "Select Barcode Date",
            value=datetime.today()
        )

    with col2:
        barcode_time = st.time_input(
            "Select Barcode Time",
            value=datetime.now().time()
        )

    # Combine date and time
    barcode_datetime = datetime.combine(barcode_date, barcode_time)

    # Apply formatted datetime
    df["Barcode Date"] = barcode_datetime.strftime("%m/%d/%Y %H:%M:%S")

    # Keep only columns that exist in df
    final_columns = [col for col in main_header if col in df.columns]
    df_final = df[final_columns]

    st.subheader("Aligned Data")
    st.dataframe(df_final.head())

    # --- Download Button ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Aligned Data")
    processed_data = output.getvalue()

    st.download_button(
        label="Download Aligned File as Excel",
        data=processed_data,
        file_name="aligned_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )