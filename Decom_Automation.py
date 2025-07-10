import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# Decom_Automation.py
def main():
    try:
        # === CONFIGURATION ===
        file_path = "Decom.xlsx"
        sheet_name = "Raw Data"
        sheet_name2 = "2025P PD24 Decom plan"
        output_sheet = "Decom Dashboard"

        # Hardcoded Decom Plan (PD24)
        #### 2/07 EDITTED To Autoamted Decom Plan(PD24)

        decom_plan={'Allianz China - Holding (CNLH)': 0,
            'Allianz China - P&C': 0,
            'Allianz Indonesia (ID)': 0,
            'Allianz Malaysia (MY)': 0,
            'Allianz Philippine - L&H (PH)': 0,
            'Allianz Singapore (AS)': 0,
            'Allianz Sri Lanka (LK)': 0,
            'Allianz Taiwan - Life (TWL)': 0,
            'Allianz Thailand (TH)': 0}
        df1 = pd.read_excel(file_path, sheet_name=sheet_name2,
                header=2, nrows=10, usecols="B:C")

        for _, row in df1.iterrows():
            key = row['Oes']       # key in dictionary
            value = row['2025 Decom Plan (PD24)']     # value to update/add
            if key in decom_plan:
                decom_plan[key] += value   # add value to existing total





        # === LOAD & FILTER RAW DATA ===
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
        df = df[['OE Name', 'Forecast End Date', 'Phase']].copy()
        df.columns = ['OE', 'Date', 'Status']
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df[df['Date'].dt.year == 2025].copy()
        df['Quarter'] = df['Date'].dt.quarter.map(lambda q: f"Q{q}") # THIS IS CORRECT OUTPUT --> CORRECT FILTER 

        # === TABLE STRUCTURE ===
        oes = list(decom_plan.keys())
        output = pd.DataFrame(index=oes)
        output["2025 Decom Plan (PD24)"] = output.index.map(decom_plan.get)

        # Completed YTD
        output["Completed YTD"] = df[df['Status'].str.lower() == "completed"].groupby("OE").size()
        output["Completed YTD"] = output["Completed YTD"].fillna(0).astype(int) # COMPLETED IS CORRECT NUMBERS 

        # In Progress (exclude completed + descoped)
        mask_in_progress = ~df["Status"].str.lower().isin(["completed", "descoped"])
        df_in_progress = df[mask_in_progress]

        output["In Progress"] = df_in_progress.groupby("OE").size()
        output["In Progress"] = output["In Progress"].fillna(0).astype(int) # IN PROGRESS IS CORRECT NUMBERS

        # Quarter-wise counts (from in-progress only)
        for q in ['Q1', 'Q2', 'Q3', 'Q4']:
            output[q] = df_in_progress[df_in_progress['Quarter'] == q].groupby('OE').size() #CORRECT NUMBERS DISPLAYED
            

        # Fill NaNs for missing quarter-OEs
        output[['Q1', 'Q2', 'Q3', 'Q4']] = output[['Q1', 'Q2', 'Q3', 'Q4']].fillna(0).astype(int)

        # Grand Total by quarter columns
        output["Grand Total"] = output[['Q1', 'Q2', 'Q3', 'Q4']].sum(axis=1)

        # YTD FC = Completed YTD + In Progress
        output["2025 YTD FC"] = output["Completed YTD"] + output["In Progress"]

        # Add Grand Total row
        # if "Grand Total" in output.index:
        #     output = output.drop("Grand Total")

        # Now safely compute grand totals
        grand_total = output.sum(numeric_only=True)
        grand_total["2025 Decom Plan (PD24)"] = sum(decom_plan.values())  # hardcoded sum = 12

        # Append grand total row
        output.loc["Grand Total"] = grand_total

        # === EXPORT TO EXCEL IN TABLE FORMAT ===
        wb = load_workbook(file_path)
        if output_sheet in wb.sheetnames:
            del wb[output_sheet]
        ws = wb.create_sheet(output_sheet)

        # Title Header
        ws["E1"] = "Planned Decom Timeline"
        ws["E1"].font = Font(bold=True)
        ws.merge_cells("E1:I1")

        # Column Headers
        headers = [
            "OE", "2025 Decom Plan (PD24)", "Completed YTD", "In Progress",
            "Q1", "Q2", "Q3", "Q4", "Grand Total", "2025 YTD FC"
        ]
        ws.append(headers)

        # Write Data
        for idx, row in output.reset_index().iterrows():
            ws.append(list(row))

        # Header Colors
        header_fill = PatternFill(start_color='FF70AD47', end_color='FF70AD47', fill_type='solid')  # Green
        data_fill = PatternFill(start_color='FFE2EFDA', end_color='FFE2EFDA', fill_type='solid') # Light Green

        # Initialise borders
        thin_border = Border(bottom=Side(style='thin', color='000000'))

        # Color in headers
        for col in range(1, len(headers) + 1):
            for row in range(1, 3):
                cell = ws[f"{get_column_letter(col)}{row}"]
                cell.fill = header_fill
                cell.font = Font(bold=True, color="FFFFFF")
            
            # Color in last row
            cell = ws[f"{get_column_letter(col)}12"]
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")

        # Color in first column
        for row in range(3,12):
            cell = ws[f"{get_column_letter(1)}{row}"]
            cell.fill = data_fill
            cell.font = Font(bold=True, color="FF000000")

        # Add borders between rows
        for col in range(1, len(headers) + 1):
            for row in range(2,13):
                cell = ws[f"{get_column_letter(col)}{row}"]
                cell.border = thin_border

        # Style Alignment
        for col in range(1, len(headers) + 1):
            for row in range(1, len(output) + 3):
                ws[f"{get_column_letter(col)}{row}"].alignment = Alignment(horizontal="center", vertical="center")

        # Autofit columns
        for col in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = max_length + 2

        # Save file
        wb.save(file_path)
        print("✅ Decom Table is completed!")
        
    except Exception as e:
        print(f"❌ Something went wrong: {e}")

if __name__ == "__main__":
    main()


