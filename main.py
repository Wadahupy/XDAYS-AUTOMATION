import streamlit as st

st.set_page_config(
    page_title="Excel Header & Consolidation Tool",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Excel Data Management App")

st.markdown(
    """
    Welcome to the **Excel Header & Consolidation Tool** 👋  
    Use the sidebar (left side) to navigate between pages:
    - 📑 **Header Alignment** – Align column headers for consistency  
    - 📂 **File Consolidation** – Combine multiple Excel files into one  
    """
)

st.info("Select a page from the sidebar to get started.")