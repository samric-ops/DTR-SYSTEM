import streamlit as st
import pandas as pd
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side
from io import BytesIO

# ---------------- PAGE CONFIG ----------------
st.set_page_config(page_title="DTR Generator (CS Form 48)", layout="wide")
st.title("📋 Daily Time Record Generator (CS Form No. 48)")

# ---------------- SIDEBAR ----------------
with st.sidebar:
    st.header("Employee Information")
    employee_name = st.text_input("Employee Name", "SAMORANOS, RICHARD P.")

    month = st.selectbox("Month", list(calendar.month_name)[1:])
    year = st.number_input("Year", min_value=2020, max_value=2100, value=2026)

    st.header("Official Office Hours")
    am_hours = st.text_input("AM Hours", "07:30 AM – 11:50 AM")
    pm_hours = st.text_input("PM Hours", "12:50 PM – 04:30 PM")
    saturday_hours = st.text_input("Saturday", "AS REQUIRED")

# ---------------- DAILY TIME INPUT ----------------
month_index = list(calendar.month_name).index(month)
num_days = calendar.monthrange(year, month_index)[1]

rows = []
for day in range(1, num_days + 1):
    weekday = calendar.weekday(year, month_index, day)

    if weekday == 5:
        rows.append({"Day": day, "AM In": "SATURDAY", "AM Out": "", "PM In": "", "PM Out": ""})
    elif weekday == 6:
        rows.append({"Day": day, "AM In": "SUNDAY", "AM Out": "", "PM In": "", "PM Out": ""})
    else:
        rows.append({
            "Day": day,
            "AM In": "07:30",
            "AM Out": "11:50",
            "PM In": "12:50",
            "PM Out": "16:30"
        })

dtr_df = pd.DataFrame(rows)

st.subheader("🕒 Daily Time Entries")
edited_df = st.data_editor(
    dtr_df,
    hide_index=True,
    use_container_width=True
)

# ---------------- GENERATE BUTTON ----------------
if st.button("📄 Generate DTR Excel File", type="primary"):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "DTR"

        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        bold = Font(bold=True)
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Set column widths
        columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
        widths = [6, 10, 10, 10, 10, 10, 10]
        for col, width in zip(columns, widths):
            ws.column_dimensions[col].width = width

        current_row = 1

        # -------- HEADER SECTION --------
        header_texts = [
            ("REPUBLIC OF THE PHILIPPINES", True),
            ("Department of Education", True),
            ("Division of Davao del Sur", True),
            ("MANUAL NATIONAL HIGH SCHOOL", True),
            ("", False),  # Empty row
            ("DAILY TIME RECORD", True),
            ("-----o0o-----", True),
            ("", False),  # Empty row
            (f"Name: {employee_name}", True),
            (f"For the month of: {month} {year}", True),
            ("", False),  # Empty row
            ("Official hours for arrival and departure", True),
            (f"Regular days: {am_hours} / {pm_hours}", True),
            (f"Saturdays: {saturday_hours}", True),
            ("", False),  # Empty row
            ("", False)   # Empty row
        ]
        
        for text, is_bold in header_texts:
            # Merge cells first
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
            # Get the TOP-LEFT cell of the merged range (this is writable)
            cell = ws.cell(row=current_row, column=1)
            cell.value = text if text is not None else ""
            cell.alignment = center
            if is_bold:
                cell.font = bold
            current_row += 1

        # -------- TABLE HEADER --------
        # Top row headers - MERGE FIRST before setting values
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 1, end_column=1)  # Day
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)  # A.M.
        ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=5)  # P.M.
        ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=7)  # Undertime

        # Set values for merged cells - ALWAYS use the TOP-LEFT cell
        ws.cell(row=current_row, column=1, value="Day").alignment = center
        ws.cell(row=current_row, column=1).font = bold
        
        ws.cell(row=current_row, column=2, value="A.M.").alignment = center
        ws.cell(row=current_row, column=2).font = bold
        
        ws.cell(row=current_row, column=4, value="P.M.").alignment = center
        ws.cell(row=current_row, column=4).font = bold
        
        ws.cell(row=current_row, column=6, value="Undertime").alignment = center
        ws.cell(row=current_row, column=6).font = bold

        # Second row sub-headers
        current_row += 1
        
        # IMPORTANT: When accessing merged cells, always use the top-left cell
        # Set sub-headers for columns 2-7 (column 1 is merged from previous row)
        sub_headers = ["", "Arrival", "Departure", "Arrival", "Departure", "Hours", "Minutes"]
        
        for col_idx in range(1, 8):  # Columns 1-7
            # For column 1, it's part of the vertical merge - DON'T overwrite it
            if col_idx == 1:
                # Just set the border for the merged cell
                ws.cell(row=current_row, column=col_idx).border = border
            else:
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = sub_headers[col_idx - 1]
                cell.alignment = center
                cell.font = bold
                cell.border = border

        current_row += 1

        # -------- TABLE DATA --------
        for _, row_data in edited_df.iterrows():
            # Day column
            day_cell = ws.cell(row=current_row, column=1)
            day_cell.value = int(row_data["Day"]) if not pd.isna(row_data["Day"]) else ""
            day_cell.alignment = center
            day_cell.border = border

            # Check if SATURDAY or SUNDAY
            if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                # Merge cells for SATURDAY/SUNDAY
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
                # Set value in the top-left cell of the merge
                merged_cell = ws.cell(row=current_row, column=2)
                merged_cell.value = str(row_data["AM In"]).strip()
                merged_cell.alignment = center
                merged_cell.border = border
                
                # Set empty values for other cells in the merged range
                for col in [3, 4, 5]:
                    cell = ws.cell(row=current_row, column=col)
                    cell.border = border
            else:
                # Regular work day
                # AM In
                am_in_cell = ws.cell(row=current_row, column=2)
                am_in_val = row_data["AM In"]
                am_in_cell.value = "" if pd.isna(am_in_val) else str(am_in_val)
                am_in_cell.alignment = center
                am_in_cell.border = border
                
                # AM Out
                am_out_cell = ws.cell(row=current_row, column=3)
                am_out_val = row_data["AM Out"]
                am_out_cell.value = "" if pd.isna(am_out_val) else str(am_out_val)
                am_out_cell.alignment = center
                am_out_cell.border = border
                
                # PM In
                pm_in_cell = ws.cell(row=current_row, column=4)
                pm_in_val = row_data["PM In"]
                pm_in_cell.value = "" if pd.isna(pm_in_val) else str(pm_in_val)
                pm_in_cell.alignment = center
                pm_in_cell.border = border
                
                # PM Out
                pm_out_cell = ws.cell(row=current_row, column=5)
                pm_out_val = row_data["PM Out"]
                pm_out_cell.value = "" if pd.isna(pm_out_val) else str(pm_out_val)
                pm_out_cell.alignment = center
                pm_out_cell.border = border

            # Undertime columns (6 and 7) - empty
            for col in [6, 7]:
                cell = ws.cell(row=current_row, column=col)
                cell.value = ""
                cell.alignment = center
                cell.border = border

            current_row += 1

        # -------- TOTAL ROW --------
        # Merge cells first
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        
        # Set value in top-left cell
        total_cell = ws.cell(row=current_row, column=1)
        total_cell.value = "TOTAL"
        total_cell.alignment = center
        total_cell.font = bold
        
        # Add border to all cells in total row
        for col in range(1, 8):
            cell = ws.cell(row=current_row, column=col)
            cell.border = border
            # Set empty values for undertime columns
            if col in [6, 7]:
                cell.value = ""

        current_row += 3

        # -------- FOOTER --------
        # Merge cells first
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 2, end_column=7)
        
        # Set value in top-left cell
        footer_cell = ws.cell(row=current_row, column=1)
        footer_cell.value = (
            "I certify on my honor that the above is a true and correct report of the\n"
            "hours of work performed, record of which was made daily at the time of\n"
            "arrival and departure from office."
        )
        footer_cell.alignment = center

        current_row += 4
        
        # Signature line - Merge first
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
        
        # Set value in top-left cell of right merge
        signature_cell = ws.cell(row=current_row, column=5)
        signature_cell.value = "Principal III"
        signature_cell.alignment = center

        # -------- PAGE SETUP --------
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.fitToHeight = 1
        ws.page_setup.fitToWidth = 1
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.75
        ws.page_margins.bottom = 0.75
        ws.page_setup.horizontalCentered = True

        # -------- SAVE AND DOWNLOAD --------
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.success("✅ DTR Excel file generated successfully!")
        
        # Create safe filename
        safe_name = "".join([c if c.isalnum() or c in "._- " else "_" for c in employee_name])
        
        st.download_button(
            "📥 Download Excel File",
            buffer.getvalue(),
            file_name=f"DTR_{safe_name}_{month}_{year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Error generating Excel file: {str(e)}")
        st.info("Please check your inputs and try again.")
