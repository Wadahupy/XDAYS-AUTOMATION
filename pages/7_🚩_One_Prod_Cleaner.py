import streamlit as st
import pandas as pd
from io import BytesIO
from msoffcrypto import OfficeFile
import io


st.set_page_config(
    page_title="One Prod Cleaner",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🚩 ONE Prod Cleaner")
st.caption("Upload an Excel file to automatically clean and standardize ONE Prod data.")
st.divider()




# === Helper: read Excel/CSV safely with automatic password bypass ===
def read_file(uploaded_file, password=None, sheet_name=0, header_row=0):
    try:
        file_ext = uploaded_file.name.split(".")[-1].lower()
        
        # Read file into bytes for reliable processing
        file_bytes = uploaded_file.read()
        file_buffer = io.BytesIO(file_bytes)

        if file_ext == "csv":
            file_buffer.seek(0)
            df = pd.read_csv(file_buffer)
            # Ensure it's a DataFrame, not a dict
            if isinstance(df, dict):
                df = pd.DataFrame(df)
            return df, None

        elif file_ext in ["xls", "xlsx"]:
            try:
                # First try to read without password
                file_buffer.seek(0)
                df = pd.read_excel(file_buffer, sheet_name=sheet_name, header=header_row, engine="openpyxl")
                # Ensure it's a DataFrame, not a dict
                if isinstance(df, dict):
                    df = pd.concat(df, ignore_index=False)
                return df, None
            except Exception as e:
                # If that fails, attempt to decrypt
                if password:
                    file_buffer.seek(0)
                    decrypted = io.BytesIO()
                    office_file = OfficeFile(file_buffer)
                    office_file.load_key(password=password)
                    office_file.decrypt(decrypted)
                    decrypted.seek(0)
                    df = pd.read_excel(decrypted, sheet_name=sheet_name, header=header_row, engine="openpyxl")
                    if isinstance(df, dict):
                        df = pd.concat(df, ignore_index=False)
                    return df, password
              
                else:
                    st.warning("⚠️ File appears to be password protected. Please enter the password below.")
                    return None, None
        else:
            st.error("Unsupported file format.")
            return None, None

    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")
        return None, None


# === Helper: Clean data by keeping only unique values per column ===
def clean_unique_values(df):
    """Remove duplicate rows while preserving headers"""
    return df.drop_duplicates().reset_index(drop=True)


# === Helper: Smart data type conversion ===
def smart_convert_types(df):
    """Intelligently convert columns to appropriate data types"""
    df_converted = df.copy()
    
    # Date columns - convert to mm/dd/yyyy format
    date_cols = [col for col in df.columns if 'date' in col.lower()]
    for col in date_cols:
        try:
            df_converted[col] = pd.to_datetime(df_converted[col], errors='coerce').dt.strftime('%m/%d/%Y')
        except:
            pass
    
    # Time columns
    time_cols = [col for col in df.columns if 'time' in col.lower()]
    for col in time_cols:
        try:
            df_converted[col] = pd.to_datetime(df_converted[col], format='%H:%M:%S', errors='coerce').dt.time
        except:
            pass
    
    # ID columns (should be string or int, not float)
    id_cols = [col for col in df.columns if 'id' in col.lower()]
    for col in id_cols:
        try:
            if df_converted[col].dtype == 'float64':
                df_converted[col] = df_converted[col].fillna(0).astype('Int64')
        except:
            pass
    
    # Amount columns (numeric)
    amount_cols = [col for col in df.columns if 'amount' in col.lower()]
    for col in amount_cols:
        try:
            df_converted[col] = pd.to_numeric(df_converted[col], errors='coerce')
        except:
            pass
    
    return df_converted

# === Helper: Data quality report ===
def generate_quality_report(df):
    """Generate data quality insights"""
    report = {}
    
    # Check for missing values
    missing = df.isnull().sum()
    report['missing_values'] = missing[missing > 0].to_dict() if missing.sum() > 0 else {}
    
    # Check for duplicate IDs
    id_cols = [col for col in df.columns if 'id' in col.lower()]
    report['duplicate_ids'] = {}
    for col in id_cols:
        dup_count = df[col].duplicated().sum()
        if dup_count > 0:
            report['duplicate_ids'][col] = int(dup_count)
    
    # Check date range
    date_cols = [col for col in df.columns if 'date' in col.lower()]
    report['date_ranges'] = {}
    for col in date_cols:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            report['date_ranges'][col] = {
                'min': str(df[col].min()),
                'max': str(df[col].max())
            }
    
    return report


# === Helper: Detect and display header information ===
def display_header_info(df):
    """Display header information and statistics"""
    st.subheader("📋 Header & Data Information")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Total Columns", len(df.columns))
    
    with col2:
        st.metric("Total Rows", len(df))
    
    with col3:
        st.metric("Unique Rows", df.drop_duplicates().shape[0])
    
    st.subheader("🔍 Column Details")
    
    header_info = pd.DataFrame({
        "Column Name": df.columns,
        "Data Type": df.dtypes.values,
        "Non-Null Count": df.notna().sum().values,
        "Unique Values": [df[col].nunique() for col in df.columns],
        "Missing Values": df.isna().sum().values
    })
    
    st.dataframe(header_info, use_container_width=True, hide_index=True)
    
    return header_info


# ---- Upload File ----
uploaded_file = st.file_uploader("Upload your CSV/Excel file", type=["csv", "xlsx"])


# ---- Main Processing Logic ----
if uploaded_file:
    st.divider()
    
    manual_password = st.text_input("Enter password (if file is encrypted)", type="password", key="manual_pwd")
    
    # Read the file with automatic password bypass
    df, used_password = read_file(uploaded_file, password=manual_password if manual_password else None)
    
    if df is not None:
        st.success("✅ File loaded successfully!")
        
        if used_password:
            st.info(f"🔓 File was protected with a password. Auto-bypass successful!")
        
        # Smart type conversion
        df = smart_convert_types(df)
        
        # Display header information
        header_info = display_header_info(df)
        
        st.divider()
        
        # Show preview of original data
        with st.expander("👁️ Original Data Preview", expanded=True):
            st.subheader("First 10 rows:")
            st.dataframe(df.head(10), use_container_width=True)
        
        st.divider()
        
        # Cleaning options
        st.subheader("🧹 Cleaning Options")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            remove_duplicates = st.checkbox("Remove Duplicate Rows", value=True)
        
        with col2:
            remove_empty = st.checkbox("Remove Empty Rows", value=False)
        
        with col3:
            remove_null_cols = st.checkbox("Remove Completely Empty Columns", value=False)
        
        # Apply cleaning
        cleaned_df = df.copy()
        
        if remove_duplicates:
            cleaned_df = clean_unique_values(cleaned_df)
        
        if remove_empty:
            cleaned_df = cleaned_df.dropna(how='all')
        
        if remove_null_cols:
            cleaned_df = cleaned_df.dropna(axis=1, how='all')
        
        st.divider()
        
        # Display cleaning results
        st.subheader("📊 Cleaning Results")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Original Rows", len(df))
        
        with col2:
            st.metric("Cleaned Rows", len(cleaned_df))
        
        with col3:
            rows_removed = len(df) - len(cleaned_df)
            st.metric("Rows Removed", rows_removed)
        
        with col4:
            reduction_pct = (rows_removed / len(df) * 100) if len(df) > 0 else 0
            st.metric("Reduction %", f"{reduction_pct:.1f}%")
        
        st.divider()
        
        # Show cleaned data preview
        with st.expander("👀 Cleaned Data Preview", expanded=True):
            st.subheader("First 10 rows of cleaned data:")
            st.dataframe(cleaned_df.head(10), use_container_width=True)
        
        st.divider()
        
        # Download cleaned file
        st.subheader("💾 Download Cleaned File")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Excel download - plain table with file-level password encryption
            output_excel = BytesIO()
            
            # Use xlsxwriter for file-level password encryption
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                cleaned_df.to_excel(writer, index=False, sheet_name='sheet1', header=True)
                
                # Add password if password was used
                if used_password or manual_password:
                    password = used_password if used_password else manual_password
                    writer.book.set_properties({'password': password})
            
            output_excel.seek(0)
            
            st.download_button(
                label="📥 Download as Excel (.xlsx)",
                data=output_excel,
                file_name=f"Cleaned_{uploaded_file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col2:
            # CSV download
            csv_output = cleaned_df.to_csv(index=False)
            st.download_button(
                label="📥 Download as CSV",
                data=csv_output,
                file_name=f"Cleaned_{uploaded_file.name.split('.')[0]}.csv",
                mime="text/csv"
            )
        
        st.divider()
        
    
else:
    st.info("👆 Please upload a CSV or Excel file to get started!")






