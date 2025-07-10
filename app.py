import streamlit as st
import Decom_Automation
import shutil

st.set_page_config(page_title="Generate Decom Report", page_icon="📊", layout="centered")
st.title("📊 Decom Dashboard Generator")
st.markdown("Upload your Excel file below and click 'Generate Report' to get the updated version.")

# Upload section
uploaded_file = st.file_uploader("📁 Upload Decom.xlsx", type=["xlsx"])

if uploaded_file is not None:
    with open("Decom.xlsx", "wb") as f:
        f.write(uploaded_file.read())
    
    st.success("✅ File uploaded successfully!")

    if st.button("Generate Report"):
        try:
            Decom_Automation.main()  # Runs and saves updated Decom.xlsx
            
            with open("Decom.xlsx", "rb") as f:
                excel_data = f.read()

            st.success("✅ Report generated! Download the updated Excel file below.")
            st.download_button(
                label="📥 Download Updated Excel File",
                data=excel_data,
                file_name="Updated_Decom.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Something went wrong:\n\n{e}")
