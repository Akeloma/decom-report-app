from openpyxl import Workbook
from flt_pvt import generate_flt_pvt_sheet
from toxic_pvt import generate_toxic_pvt_sheet
from Group_FLT_Details import generate_group_flt_details
from Group_Toxic_Details import generate_group_toxic_details
from Local_FLT_Details import generate_local_flt_details
from Local_Toxic_Details import generate_local_toxic_details

def generate_full_report():
    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    generate_flt_pvt_sheet(wb)
    generate_toxic_pvt_sheet(wb)
    generate_group_flt_details(wb)
    generate_group_toxic_details(wb)
    generate_local_flt_details(wb)
    generate_local_toxic_details(wb)

    wb.save("Archer_Toxic_Report_Final.xlsx")

if __name__ == "__main__":
    generate_full_report()
