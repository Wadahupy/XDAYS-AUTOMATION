import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="CMS ↔ TEXXEN Alignment", layout="wide")
st.title("🔄 CMS ↔ TEXXEN Header Alignment")

# ---------------------------
# Alignment mappings
# ---------------------------
CMS_TO_TEXXEN_MAP = {
    "_id": "_id",
    "debtorId": "Debtor ID",
    "accountNumber": "Account No.",
    "cardNumber": "Card No.",
    "chCode": "Card No.",
    "accountName": "Debtor",
    "OB": "Balance",
    "principal": "principal",
    "endoDate": "endoDate",
    "bankName": "Client",
    "placement": "placement",
    "productType": "Product Type",
    "cycle": "Cycle",
    "dpd": "dpd",
    "level": "level",
    "contactSource": "contactSource",
    "status": "Status1",
    "substatus": "Status2",
    "groupStatus": "groupStatus",
    "dispositionSource": "dispositionSource",
    "ptpAmount": "PTP Amount",
    "paymentAmount": "Claim Paid Amount",
    "startDate": "PTP Date1",
    "endDate": "PTP Date2",
    "orNumber": "orNumber",
    "rfd": "Reason For Default",
    "newEmail": "newEmail",
    "newContact": "newContact",
    "newAddress": "newAddress",
    "notes": "Remark",
    "dispositionType": "Remark Type",
    "dialedNumber": "Dialed Number",
    "contactType": "Contact Type",
    "autoStat": "autoStat",
    "viciRecordings": "viciRecordings",
    "duration": "Talk Time Duration",
    "agent": "Remark By",
    "barcodeDate": "Date1",
    "time": "Time",
    "Created Date": "Date2",
}

TEXXEN_TO_CMS_MAP = {
    "S.No": "S.No",
    "Date": "barcodeDate",
    "Time": "time",
    "Debtor": "accountName",
    "Account No.": "accountNumber",
    "Card No.": "chCode",
    "Service No.": "Service No.",
    "DPD": "dpd",
    "Call Status": "Call Status",
    "Status": "status",
    "Remark": "notes",
    "Remark By": "agent",
    "Remark Type": "dispositionType",
    "Client": "bankName",
    "Product Type": "productType",
    "PTP Amount": "ptpAmount",
    "PTP Date": "startDate",
    "Claim Paid Amount": "paymentAmount",
    "Dialed Number": "dialedNumber",
    "Balance": "OB",
    "Contact Type": "contactType",
    "Debtor ID": "debtorId",
}

# ---------------------------
# Data alignment function
# ---------------------------
def align_dataframe(df, mapping, direction):
    aligned_data = {}
    
    df.columns = df.columns.str.strip()
    col_lookup = {col.lower().strip(): col for col in df.columns}
    
    # Critical columns should only match exactly
    exact_match_only = {"debtorid", "accountname"}

    for source_col, target_col in mapping.items():
        source_clean = source_col.lower().strip()
        matching_col = None
        
        # Exact match first
        if source_clean in col_lookup:
            matching_col = col_lookup[source_clean]
        # Match without spaces or periods
        else:
            source_no_spaces = source_clean.replace(" ", "").replace(".", "")
            for col_clean, col_original in col_lookup.items():
                col_no_spaces = col_clean.replace(" ", "").replace(".", "")
                if source_no_spaces == col_no_spaces:
                    matching_col = col_original
                    break
        # Partial match (skip critical columns)
        if not matching_col and source_clean not in exact_match_only:
            for col_original in df.columns:
                if source_clean in col_original.lower() or col_original.lower() in source_clean:
                    if len(source_clean) > 3:
                        matching_col = col_original
                        break
        
        aligned_data[target_col] = df[matching_col] if matching_col else ""
    
    aligned_df = pd.DataFrame(aligned_data)
    
    # CMS → TEXXEN special handling
    if direction == "CMS → TEXXEN":
        if "Status1" in aligned_df.columns and "Status2" in aligned_df.columns:
            status_split = aligned_df["Status1"].astype(str).str.split("-", n=1, expand=True)
            aligned_df["Status1"] = status_split[0].str.strip()
            aligned_df["Status2"] = status_split[1].str.strip() if status_split.shape[1] > 1 else ""
        if "Date1" in aligned_df.columns and "Date2" in aligned_df.columns:
            aligned_df["Date2"] = aligned_df["Date1"]
        if "PTP Date1" in aligned_df.columns and "PTP Date2" in aligned_df.columns:
            aligned_df["PTP Date2"] = aligned_df["PTP Date1"]
    
    # TEXXEN → CMS special handling
    elif direction == "TEXXEN → CMS":
        substatus_col = next((c for c in df.columns if "substatus" in c.lower()), None)
        if "status" in aligned_df.columns and substatus_col:
            aligned_data_dict = aligned_df.to_dict('list')
            concatenated_status = [
                f"{s} - {sub}" if str(sub).strip() else str(s)
                for s, sub in zip(aligned_data_dict["status"], df[substatus_col])
            ]
            aligned_df["status"] = concatenated_status
        if "barcodeDate" in aligned_df.columns:
            aligned_df["barcodeDate"] = pd.to_datetime(aligned_df["barcodeDate"], errors='coerce').dt.date

    return aligned_df

# ---------------------------
# Streamlit interface
# ---------------------------
st.subheader("⚙️ Align Your File")
alignment_type = st.selectbox("Select Alignment Direction", ["CMS → TEXXEN", "TEXXEN → CMS"])
uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

if uploaded_file and alignment_type:
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        st.info(f"✅ File loaded successfully! Rows: {len(df)}, Columns: {len(df.columns)}")
        st.write("**📋 Your file headers:**", ", ".join(df.columns.tolist()))
        
        mapping = CMS_TO_TEXXEN_MAP if alignment_type == "CMS → TEXXEN" else TEXXEN_TO_CMS_MAP
        direction = "CMS → TEXXEN" if alignment_type == "CMS → TEXXEN" else "TEXXEN → CMS"
        aligned_df = align_dataframe(df, mapping, direction)
        
        st.dataframe(aligned_df.head(min(100, len(aligned_df))), use_container_width=True)
    
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
else:
    st.warning("⚠️ Please select alignment direction and upload a file to proceed.")