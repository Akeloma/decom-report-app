import streamlit as st
import Decom_Automation
import toxic_flt_table

# === Set page config ===
st.set_page_config(page_title="Report Generator", page_icon="📊", layout="centered")

# === Sidebar navigation ===
page = st.sidebar.selectbox(
    "🔧 Select Report Type",
    ["Decom Automation", "Toxic & FLT Table"]
)

st.title("📊 IT Governance Automation Portal")

# === Shared uploader helper ===
def save_uploaded_file(uploaded_file, filename):
    with open(filename, "wb") as f:
        f.write(uploaded_file.read())
    return filename

# === Page 1: Decom Automation ===
if page == "Decom Automation":
    st.subheader("📘 Decom Automation Tool")
    st.markdown("Upload your **Decom.xlsx** file below and click Generate Report.")

    uploaded_file = st.file_uploader("📁 Upload Decom.xlsx", type=["xlsx"], key="decom")

    if uploaded_file:
        save_uploaded_file(uploaded_file, "Decom.xlsx")
        st.success("✅ File uploaded successfully.")

        if st.button("🧾 Generate Decom Report"):
            try:
                Decom_Automation.main()

                with open("Decom.xlsx", "rb") as f:
                    excel_data = f.read()

                st.success("✅ Report generated! Download below:")
                st.download_button(
                    label="📥 Download Updated Decom Report",
                    data=excel_data,
                    file_name="Updated_Decom.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Something went wrong:\n\n{e}")

# === Page 2: Toxic & FLT Table ===
elif page == "Toxic & FLT Table":
    st.subheader("🧪 Toxic & FLT Table Generator")
    st.markdown("Upload your **manual calculated.xlsx** file below and click Generate Report.")

    uploaded_file = st.file_uploader("📁 Upload manual calculated.xlsx", type=["xlsx"], key="toxic")

    if uploaded_file:
        save_uploaded_file(uploaded_file, "manual calculated.xlsx")
        st.success("✅ File uploaded successfully.")

        if st.button("🧠 Generate Toxic & FLT Report"):
            try:
                toxic_flt_table.main()

                with open("Detox_Tables_Formatted.xlsx", "rb") as f:
                    excel_data = f.read()

                st.success("✅ Toxic & FLT Report generated! Download below:")
                st.download_button(
                    label="📥 Download Detox Report",
                    data=excel_data,
                    file_name="Detox_Tables_Formatted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Something went wrong:\n\n{e}")
