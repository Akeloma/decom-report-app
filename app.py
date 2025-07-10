import streamlit as st
import Decom_Automation

# Set page config
st.set_page_config(page_title="Generate Decom Report", page_icon="📊", layout="centered")

# Title
st.title("📊 Decom Dashboard Generator")

# Instructions
st.markdown("Click the button below to generate your latest Decom Table in Excel.")

# Button
if st.button("Generate Report"):
    try:
        Decom_Automation.main()
        st.success("✅ Report generated successfully! Check your Excel file.")
    except Exception as e:
        st.error(f"❌ Something went wrong:\n\n{e}")
