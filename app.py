import streamlit as st
import Decom_Automation

st.set_page_config(page_title="Generate Decom Report", page_icon="📊", layout="centered")

st.title("📊 Decom Dashboard Generator")
st.markdown("Click the button below to generate your latest Decom Table in Excel.")

if st.button("Generate Report"):
    try:
        Decom_Automation.main()  # This should save 'Decom.xlsx'
        
        # Read the updated file back in binary mode
        with open("Decom.xlsx", "rb") as f:
            excel_data = f.read()
        
        st.success("✅ Report generated successfully!")
        st.download_button(
            label="📥 Download Updated Excel File",
            data=excel_data,
            file_name="Updated_Decom.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Something went wrong:\n\n{e}")
