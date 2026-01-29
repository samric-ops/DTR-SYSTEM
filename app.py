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
    employee_no = st.text_input("Employee Number", "7220970")

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
        left = Alignment(horizontal="left", vertical="center")
        right = Alignment(horizontal="right", vertical="center")
        bold = Font(bold=True)
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        
        # Use smaller font for compact design
        small_font = Font(size=9)
        header_font = Font(bold=True, size=10)

        # Set narrower column widths for half-page format
        # Total width for 7 columns should fit half of A4
        widths = [4, 8, 8, 8, 8, 8, 8]  # Narrower columns
        for i, w in enumerate(widths, 1):
            col_letter = chr(64 + i)
            ws.column_dimensions[col_letter].width = w

        # -------- FIRST DTR (LEFT SIDE) --------
        start_row = 1
        
        # HEADER FOR FIRST DTR
        # Line 1: REPUBLIC OF THE PHILIPPINES (centered)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "REPUBLIC OF THE PHILIPPINES"
        cell.alignment = center
        cell.font = bold
        start_row += 1
        
        # Line 2: Department of Education (centered)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "Department of Education"
        cell.alignment = center
        cell.font = bold
        start_row += 1
        
        # Line 3: Division of Davao del Sur (centered)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "Division of Davao del Sur"
        cell.alignment = center
        cell.font = bold
        start_row += 1
        
        # Line 4: Manual National High School (centered)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "Manual National High School"
        cell.alignment = center
        cell.font = bold
        start_row += 2  # Extra space
        
        # Line with Civil Service Form and Employee No.
        # Left side: Civil Service Form No. 48
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "Civil Service Form No. 48"
        cell.alignment = left
        cell.font = bold
        
        # Right side: Employee No. and number
        ws.merge_cells(start_row=start_row, start_column=5, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=5)
        cell.value = f"Employee No.    {employee_no}"
        cell.alignment = right
        cell.font = bold
        start_row += 2
        
        # DAILY TIME RECORD (centered)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "DAILY TIME RECORD"
        cell.alignment = center
        cell.font = Font(bold=True, size=12)
        start_row += 1
        
        # ---o0o--- line
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "---o0o---"
        cell.alignment = center
        cell.font = bold
        start_row += 2
        
        # Employee Name
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = employee_name
        cell.alignment = center
        cell.font = bold
        start_row += 1
        
        # "(Name)" label
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "(Name)"
        cell.alignment = center
        cell.font = small_font
        start_row += 1
        
        # "For the month of" with month and year
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = f"For the month of __________ {month} __________ {year}"
        cell.alignment = center
        cell.font = small_font
        start_row += 2
        
        # Official hours section
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "Official hours for arrival and departure"
        cell.alignment = center
        cell.font = small_font
        start_row += 1
        
        # Regular days hours
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = f"Regular days: {am_hours} / {pm_hours}"
        cell.alignment = center
        cell.font = small_font
        start_row += 1
        
        # Saturdays
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = f"Saturdays: {saturday_hours}"
        cell.alignment = center
        cell.font = small_font
        start_row += 2
        
        # -------- TABLE HEADER FOR FIRST DTR --------
        # Top row headers
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + 1, end_column=1)  # Day
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=3)  # A.M.
        ws.merge_cells(start_row=start_row, start_column=4, end_row=start_row, end_column=5)  # P.M.
        ws.merge_cells(start_row=start_row, start_column=6, end_row=start_row, end_column=7)  # Undertime
        
        # Main headers
        ws.cell(row=start_row, column=1, value="Day").alignment = center
        ws.cell(row=start_row, column=1).font = header_font
        
        ws.cell(row=start_row, column=2, value="A.M.").alignment = center
        ws.cell(row=start_row, column=2).font = header_font
        
        ws.cell(row=start_row, column=4, value="P.M.").alignment = center
        ws.cell(row=start_row, column=4).font = header_font
        
        ws.cell(row=start_row, column=6, value="Undertime").alignment = center
        ws.cell(row=start_row, column=6).font = header_font
        
        # Second row sub-headers
        start_row += 1
        
        sub_headers = ["", "Arrival", "Departure", "Arrival", "Departure", "Hours", "Minutes"]
        for col_idx in range(1, 8):
            cell = ws.cell(row=start_row, column=col_idx)
            cell.value = sub_headers[col_idx - 1]
            cell.alignment = center
            cell.font = Font(bold=True, size=8)
            cell.border = border
        
        start_row += 1
        
        # -------- TABLE DATA FOR FIRST DTR (First 15 days) --------
        first_half = edited_df.iloc[:15] if len(edited_df) > 15 else edited_df
        
        for _, row_data in first_half.iterrows():
            # Day column
            day_cell = ws.cell(row=start_row, column=1)
            day_val = row_data["Day"]
            day_cell.value = int(day_val) if not pd.isna(day_val) else ""
            day_cell.alignment = center
            day_cell.border = border
            day_cell.font = small_font

            if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=5)
                merged_cell = ws.cell(row=start_row, column=2)
                merged_cell.value = str(row_data["AM In"]).strip()
                merged_cell.alignment = center
                merged_cell.border = border
                merged_cell.font = small_font
                
                for col in [3, 4, 5]:
                    cell = ws.cell(row=start_row, column=col)
                    cell.border = border
            else:
                for col_idx, col_name in [(2, "AM In"), (3, "AM Out"), (4, "PM In"), (5, "PM Out")]:
                    cell = ws.cell(row=start_row, column=col_idx)
                    val = row_data[col_name]
                    cell.value = "" if pd.isna(val) else str(val)
                    cell.alignment = center
                    cell.border = border
                    cell.font = small_font

            # Undertime columns
            for col in [6, 7]:
                cell = ws.cell(row=start_row, column=col)
                cell.value = ""
                cell.alignment = center
                cell.border = border
                cell.font = small_font

            start_row += 1
        
        # Add remaining rows if less than 15 days
        for _ in range(len(first_half), 15):
            for col in range(1, 8):
                cell = ws.cell(row=start_row, column=col)
                cell.value = ""
                cell.border = border
                cell.font = small_font
            start_row += 1
        
        # TOTAL row for first DTR
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)
        total_cell = ws.cell(row=start_row, column=1)
        total_cell.value = "TOTAL"
        total_cell.alignment = center
        total_cell.font = bold
        total_cell.border = border
        
        for col in [6, 7]:
            cell = ws.cell(row=start_row, column=col)
            cell.value = ""
            cell.border = border
        
        start_row += 4  # Space for second DTR
        
        # -------- SECOND DTR (RIGHT SIDE - SAME FORMAT) --------
        # Repeat the same structure but at column H onward (for second half of page)
        # For simplicity, we'll create it below the first one
        
        # Add separator line
        for col in range(1, 8):
            ws.cell(row=start_row, column=col).value = "─" * 15
        
        start_row += 2
        
        # HEADER FOR SECOND DTR
        # Line 1: REPUBLIC OF THE PHILIPPINES
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "REPUBLIC OF THE PHILIPPINES"
        cell.alignment = center
        cell.font = bold
        start_row += 1
        
        # Line 2-4: Same as first DTR
        for text in ["Department of Education", "Division of Davao del Sur", "Manual National High School"]:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
            cell = ws.cell(row=start_row, column=1)
            cell.value = text
            cell.alignment = center
            cell.font = bold
            start_row += 1
        
        start_row += 1  # Extra space
        
        # Civil Service Form and Employee No.
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "Civil Service Form No. 48"
        cell.alignment = left
        cell.font = bold
        
        ws.merge_cells(start_row=start_row, start_column=5, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=5)
        cell.value = f"Employee No.    {employee_no}"
        cell.alignment = right
        cell.font = bold
        start_row += 2
        
        # DAILY TIME RECORD
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "DAILY TIME RECORD"
        cell.alignment = center
        cell.font = Font(bold=True, size=12)
        start_row += 1
        
        # ---o0o--- line
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "---o0o---"
        cell.alignment = center
        cell.font = bold
        start_row += 2
        
        # Employee Name
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = employee_name
        cell.alignment = center
        cell.font = bold
        start_row += 1
        
        # "(Name)" label
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "(Name)"
        cell.alignment = center
        cell.font = small_font
        start_row += 1
        
        # "For the month of"
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = f"For the month of __________ {month} __________ {year}"
        cell.alignment = center
        cell.font = small_font
        start_row += 2
        
        # Official hours
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "Official hours for arrival and departure"
        cell.alignment = center
        cell.font = small_font
        start_row += 1
        
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = f"Regular days: {am_hours} / {pm_hours}"
        cell.alignment = center
        cell.font = small_font
        start_row += 1
        
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = f"Saturdays: {saturday_hours}"
        cell.alignment = center
        cell.font = small_font
        start_row += 2
        
        # -------- TABLE HEADER FOR SECOND DTR --------
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + 1, end_column=1)
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=3)
        ws.merge_cells(start_row=start_row, start_column=4, end_row=start_row, end_column=5)
        ws.merge_cells(start_row=start_row, start_column=6, end_row=start_row, end_column=7)
        
        ws.cell(row=start_row, column=1, value="Day").alignment = center
        ws.cell(row=start_row, column=1).font = header_font
        
        ws.cell(row=start_row, column=2, value="A.M.").alignment = center
        ws.cell(row=start_row, column=2).font = header_font
        
        ws.cell(row=start_row, column=4, value="P.M.").alignment = center
        ws.cell(row=start_row, column=4).font = header_font
        
        ws.cell(row=start_row, column=6, value="Undertime").alignment = center
        ws.cell(row=start_row, column=6).font = header_font
        
        start_row += 1
        
        for col_idx in range(1, 8):
            cell = ws.cell(row=start_row, column=col_idx)
            cell.value = sub_headers[col_idx - 1]
            cell.alignment = center
            cell.font = Font(bold=True, size=8)
            cell.border = border
        
        start_row += 1
        
        # -------- TABLE DATA FOR SECOND DTR (Days 16-31) --------
        second_half = edited_df.iloc[15:] if len(edited_df) > 15 else pd.DataFrame()
        
        if len(second_half) > 0:
            for _, row_data in second_half.iterrows():
                day_cell = ws.cell(row=start_row, column=1)
                day_val = row_data["Day"]
                day_cell.value = int(day_val) if not pd.isna(day_val) else ""
                day_cell.alignment = center
                day_cell.border = border
                day_cell.font = small_font

                if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                    ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=5)
                    merged_cell = ws.cell(row=start_row, column=2)
                    merged_cell.value = str(row_data["AM In"]).strip()
                    merged_cell.alignment = center
                    merged_cell.border = border
                    merged_cell.font = small_font
                else:
                    for col_idx, col_name in [(2, "AM In"), (3, "AM Out"), (4, "PM In"), (5, "PM Out")]:
                        cell = ws.cell(row=start_row, column=col_idx)
                        val = row_data[col_name]
                        cell.value = "" if pd.isna(val) else str(val)
                        cell.alignment = center
                        cell.border = border
                        cell.font = small_font

                for col in [6, 7]:
                    cell = ws.cell(row=start_row, column=col)
                    cell.value = ""
                    cell.alignment = center
                    cell.border = border
                    cell.font = small_font

                start_row += 1
        
        # Fill remaining rows
        rows_filled = len(second_half)
        for _ in range(rows_filled, 15):
            for col in range(1, 8):
                cell = ws.cell(row=start_row, column=col)
                cell.value = ""
                cell.border = border
                cell.font = small_font
            start_row += 1
        
        # TOTAL row for second DTR
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)
        total_cell = ws.cell(row=start_row, column=1)
        total_cell.value = "TOTAL"
        total_cell.alignment = center
        total_cell.font = bold
        total_cell.border = border
        
        for col in [6, 7]:
            cell = ws.cell(row=start_row, column=col)
            cell.value = ""
            cell.border = border
        
        start_row += 4
        
        # -------- FOOTER (CERTIFICATION) --------
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + 2, end_column=7)
        footer_cell = ws.cell(row=start_row, column=1)
        footer_cell.value = (
            "I certify on my honor that the above is a true and correct report of the\n"
            "hours of work performed, record of which was made daily at the time of\n"
            "arrival and departure from office."
        )
        footer_cell.alignment = center
        footer_cell.font = small_font

        start_row += 4
        
        # Signature line
        ws.merge_cells(start_row=start_row, start_column=5, end_row=start_row, end_column=7)
        signature_cell = ws.cell(row=start_row, column=5)
        signature_cell.value = "Principal III"
        signature_cell.alignment = center
        signature_cell.font = small_font

        # -------- PAGE SETUP FOR HALF-PAGE --------
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # Landscape for two DTRs side by side
        ws.page_setup.fitToHeight = 1
        ws.page_setup.fitToWidth = 1
        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5
        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = False
        
        # Set print area for two DTRs per page
        ws.print_area = f'A1:G{start_row}'

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
            file_name=f"DTR_CSForm48_{safe_name}_{month}_{year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Error generating Excel file: {str(e)}")
        st.info("Please check your inputs and try again.")
