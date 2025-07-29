import streamlit as st
import Decom_Automation
import toxic_flt_table
import run_all_reports
import amendedToxicFLT


# === Set page config ===
st.set_page_config(page_title="Report Generator", page_icon="📊", layout="centered")

# === Sidebar navigation ===
page = st.sidebar.selectbox(
    "🔧 Select Report Type",
    ["Decom Automation", "Toxic & FLT Table", "One-Click Full Toxic & FLT", "Desmond's Pivot Tables", "Amended Toxic & FLT"]
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

                with open("Toxic&FLT_Tables.xlsx", "rb") as f:
                    excel_data = f.read()

                st.success("✅ Toxic & FLT Report generated! Download below:")
                st.download_button(
                    label="📥 Download Toxic & FLT Tables",
                    data=excel_data,
                    file_name="Toxic&FLT_Tables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Something went wrong:\n\n{e}")

# === Page 3: One-Click Full Toxic & FLT ===
elif page == "One-Click Full Toxic & FLT":
    st.subheader("🧩 One-Click Full Toxic & FLT Automation")
    st.markdown("Upload your **manual calculated.xlsx** file and run all four reports with a single click.")

    uploaded_file = st.file_uploader("📁 Upload manual calculated.xlsx", type=["xlsx"], key="oneclick")

    if uploaded_file:
        save_uploaded_file(uploaded_file, "manual calculated.xlsx")
        st.success("✅ File uploaded successfully.")

        if st.button("🚀 Run All Reports"):
            try:
                run_all_TF.run_all()

                with open("manual calculated.xlsx", "rb") as f:
                    excel_data = f.read()
                
                st.success("✅ All 4 reports generated in one file! Download below:")
                st.download_button(
                    label="📥 Download Updated Toxic & FLT Report",
                    data=excel_data,
                    file_name="Updated_Toxic_FLT_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


            except Exception as e:
                st.error(f"❌ Error occurred:\n\n{e}")

# === Page 4: Desmond's Pivot Table ===
elif page == "Desmond's Pivot Tables":
    st.subheader("🖥️ Automated Pivot Tables (Toxic & FLT) for Desmond")
    st.markdown("Upload any Excel file and generate the full Archer report.")

    uploaded_file = st.file_uploader("📁 Upload your Excel file", type=["xlsx"], key="archer")

    if uploaded_file:
        input_filename = uploaded_file.name
        with open(input_filename, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"✅ Uploaded file: {input_filename}")

        if st.button("🧠 Generate Archer Report"):
            try:
                run_all_reports.generate_full_report()
                st.success("✅ Report generated successfully!")

                with open("Archer_Toxic_Report_Final.xlsx", "rb") as file:
                    st.download_button(
                        label="📥 Download Excel Report",
                        data=file,
                        file_name="Archer_Toxic_Report_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ Error occurred during generation:\n\n{e}")

# === Page 5: Amended Toxic & FLT Table ===
elif page == "Amended Toxic & FLT":
    st.subheader("♻️ Amended Toxic & FLT Report")
    st.markdown("Upload your file below. This version removes quarter columns and adds the Year column.")

    uploaded_file = st.file_uploader("📁 Upload yourfile.xlsx", type=["xlsx"], key="amended")

    if uploaded_file:
        save_uploaded_file(uploaded_file, "manual calculated.xlsx")
        st.success("✅ File uploaded successfully.")

        if st.button("📊 Generate Amended Report"):
            try:
                amendedToxicFLT.main()

                with open("Toxic&FLT_Tables.xlsx", "rb") as file:
                    st.download_button(
                        label="📥 Download Amended Toxic & FLT Report",
                        data=file,
                        file_name="Amended_Toxic_FLT_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ Error occurred:\n\n{e}")
