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
        
        small_font = Font(size=9)
        header_font = Font(bold=True, size=10)

        # Set column widths
        widths = [4, 8, 8, 8, 8, 8, 8]
        for i, w in enumerate(widths, 1):
            col_letter = chr(64 + i)
            ws.column_dimensions[col_letter].width = w

        current_row = 1

        # -------- FIRST DTR (TOP HALF) --------
        
        # HEADER - SET VALUES FIRST, THEN MERGE
        # Line 1
        cell = ws.cell(row=current_row, column=1, value="REPUBLIC OF THE PHILIPPINES")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 1
        
        # Line 2
        cell = ws.cell(row=current_row, column=1, value="Department of Education")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 1
        
        # Line 3
        cell = ws.cell(row=current_row, column=1, value="Division of Davao del Sur")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 1
        
        # Line 4
        cell = ws.cell(row=current_row, column=1, value="Manual National High School")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 2
        
        # Civil Service Form and Employee No. - SET BOTH FIRST
        # Left side
        cell_left = ws.cell(row=current_row, column=1, value="Civil Service Form No. 48")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        cell_left.alignment = left
        cell_left.font = bold
        
        # Right side
        cell_right = ws.cell(row=current_row, column=5, value=f"Employee No.    {employee_no}")
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
        cell_right.alignment = right
        cell_right.font = bold
        current_row += 2
        
        # DAILY TIME RECORD
        cell = ws.cell(row=current_row, column=1, value="DAILY TIME RECORD")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = Font(bold=True, size=12)
        current_row += 1
        
        # ---o0o--- line
        cell = ws.cell(row=current_row, column=1, value="---o0o---")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 2
        
        # Employee Name
        cell = ws.cell(row=current_row, column=1, value=employee_name)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 1
        
        # "(Name)" label
        cell = ws.cell(row=current_row, column=1, value="(Name)")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 1
        
        # "For the month of"
        cell = ws.cell(row=current_row, column=1, value=f"For the month of __________ {month} __________ {year}")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 2
        
        # Official hours section
        cell = ws.cell(row=current_row, column=1, value="Official hours for arrival and departure")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 1
        
        # Regular days hours
        cell = ws.cell(row=current_row, column=1, value=f"Regular days: {am_hours} / {pm_hours}")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 1
        
        # Saturdays
        cell = ws.cell(row=current_row, column=1, value=f"Saturdays: {saturday_hours}")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 2
        
        # -------- TABLE HEADER FOR FIRST DTR --------
        # IMPORTANT: Set values in TOP-LEFT cells BEFORE merging
        
        # Day header (will be merged vertically)
        day_cell = ws.cell(row=current_row, column=1, value="Day")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 1, end_column=1)
        day_cell.alignment = center
        day_cell.font = header_font
        
        # A.M. header (will be merged horizontally)
        am_cell = ws.cell(row=current_row, column=2, value="A.M.")
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
        am_cell.alignment = center
        am_cell.font = header_font
        
        # P.M. header (will be merged horizontally)
        pm_cell = ws.cell(row=current_row, column=4, value="P.M.")
        ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=5)
        pm_cell.alignment = center
        pm_cell.font = header_font
        
        # Undertime header (will be merged horizontally)
        under_cell = ws.cell(row=current_row, column=6, value="Undertime")
        ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=7)
        under_cell.alignment = center
        under_cell.font = header_font
        
        # Second row sub-headers
        current_row += 1
        
        sub_headers = ["", "Arrival", "Departure", "Arrival", "Departure", "Hours", "Minutes"]
        for col_idx in range(1, 8):
            # IMPORTANT: Don't touch column 1 - it's part of vertical merge
            if col_idx != 1:
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = sub_headers[col_idx - 1]
                cell.alignment = center
                cell.font = Font(bold=True, size=8)
                cell.border = border
            else:
                # Just add border to column 1 (already has "Day" value from merged cell)
                ws.cell(row=current_row, column=1).border = border
        
        current_row += 1
        
        # -------- TABLE DATA FOR FIRST DTR (First 15 days) --------
        first_half = edited_df.iloc[:15] if len(edited_df) > 15 else edited_df
        
        for _, row_data in first_half.iterrows():
            # Day column
            day_cell = ws.cell(row=current_row, column=1)
            day_val = row_data["Day"]
            day_cell.value = int(day_val) if not pd.isna(day_val) else ""
            day_cell.alignment = center
            day_cell.border = border
            day_cell.font = small_font

            if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                # SET VALUE FIRST, then merge
                sat_cell = ws.cell(row=current_row, column=2, value=str(row_data["AM In"]).strip())
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
                sat_cell.alignment = center
                sat_cell.border = border
                sat_cell.font = small_font
                
                # Add borders to other cells in merged range
                for col in [3, 4, 5]:
                    ws.cell(row=current_row, column=col).border = border
            else:
                for col_idx, col_name in [(2, "AM In"), (3, "AM Out"), (4, "PM In"), (5, "PM Out")]:
                    cell = ws.cell(row=current_row, column=col_idx)
                    val = row_data[col_name]
                    cell.value = "" if pd.isna(val) else str(val)
                    cell.alignment = center
                    cell.border = border
                    cell.font = small_font

            # Undertime columns
            for col in [6, 7]:
                cell = ws.cell(row=current_row, column=col)
                cell.value = ""
                cell.alignment = center
                cell.border = border
                cell.font = small_font

            current_row += 1
        
        # Add empty rows if less than 15 days
        for _ in range(len(first_half), 15):
            for col in range(1, 8):
                cell = ws.cell(row=current_row, column=col)
                cell.value = ""
                cell.border = border
                cell.font = small_font
            current_row += 1
        
        # TOTAL row for first DTR
        total_cell = ws.cell(row=current_row, column=1, value="TOTAL")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        total_cell.alignment = center
        total_cell.font = bold
        total_cell.border = border
        
        for col in [6, 7]:
            cell = ws.cell(row=current_row, column=col)
            cell.value = ""
            cell.border = border
        
        current_row += 4  # Space between DTRs
        
        # -------- SEPARATOR LINE --------
        for col in range(1, 8):
            ws.cell(row=current_row, column=col, value="─" * 15)
        current_row += 2
        
        # -------- SECOND DTR (BOTTOM HALF) --------
        # HEADER FOR SECOND DTR - SAME PATTERN: SET VALUE FIRST, THEN MERGE
        
        # Line 1
        cell = ws.cell(row=current_row, column=1, value="REPUBLIC OF THE PHILIPPINES")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 1
        
        # Lines 2-4
        for text in ["Department of Education", "Division of Davao del Sur", "Manual National High School"]:
            cell = ws.cell(row=current_row, column=1, value=text)
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
            cell.alignment = center
            cell.font = bold
            current_row += 1
        
        current_row += 1
        
        # Civil Service Form and Employee No.
        cell_left = ws.cell(row=current_row, column=1, value="Civil Service Form No. 48")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        cell_left.alignment = left
        cell_left.font = bold
        
        cell_right = ws.cell(row=current_row, column=5, value=f"Employee No.    {employee_no}")
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
        cell_right.alignment = right
        cell_right.font = bold
        current_row += 2
        
        # DAILY TIME RECORD
        cell = ws.cell(row=current_row, column=1, value="DAILY TIME RECORD")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = Font(bold=True, size=12)
        current_row += 1
        
        # ---o0o--- line
        cell = ws.cell(row=current_row, column=1, value="---o0o---")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 2
        
        # Employee Name
        cell = ws.cell(row=current_row, column=1, value=employee_name)
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = bold
        current_row += 1
        
        # "(Name)" label
        cell = ws.cell(row=current_row, column=1, value="(Name)")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 1
        
        # "For the month of"
        cell = ws.cell(row=current_row, column=1, value=f"For the month of __________ {month} __________ {year}")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 2
        
        # Official hours
        cell = ws.cell(row=current_row, column=1, value="Official hours for arrival and departure")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 1
        
        cell = ws.cell(row=current_row, column=1, value=f"Regular days: {am_hours} / {pm_hours}")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 1
        
        cell = ws.cell(row=current_row, column=1, value=f"Saturdays: {saturday_hours}")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        cell.alignment = center
        cell.font = small_font
        current_row += 2
        
        # -------- TABLE HEADER FOR SECOND DTR --------
        # SET VALUES FIRST, THEN MERGE
        day_cell = ws.cell(row=current_row, column=1, value="Day")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 1, end_column=1)
        day_cell.alignment = center
        day_cell.font = header_font
        
        am_cell = ws.cell(row=current_row, column=2, value="A.M.")
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
        am_cell.alignment = center
        am_cell.font = header_font
        
        pm_cell = ws.cell(row=current_row, column=4, value="P.M.")
        ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=5)
        pm_cell.alignment = center
        pm_cell.font = header_font
        
        under_cell = ws.cell(row=current_row, column=6, value="Undertime")
        ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=7)
        under_cell.alignment = center
        under_cell.font = header_font
        
        current_row += 1
        
        # Sub-headers
        for col_idx in range(1, 8):
            if col_idx != 1:
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = sub_headers[col_idx - 1]
                cell.alignment = center
                cell.font = Font(bold=True, size=8)
                cell.border = border
            else:
                ws.cell(row=current_row, column=1).border = border
        
        current_row += 1
        
        # -------- TABLE DATA FOR SECOND DTR (Days 16-31) --------
        second_half = edited_df.iloc[15:] if len(edited_df) > 15 else pd.DataFrame()
        
        if len(second_half) > 0:
            for _, row_data in second_half.iterrows():
                day_cell = ws.cell(row=current_row, column=1)
                day_val = row_data["Day"]
                day_cell.value = int(day_val) if not pd.isna(day_val) else ""
                day_cell.alignment = center
                day_cell.border = border
                day_cell.font = small_font

                if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                    sat_cell = ws.cell(row=current_row, column=2, value=str(row_data["AM In"]).strip())
                    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
                    sat_cell.alignment = center
                    sat_cell.border = border
                    sat_cell.font = small_font
                else:
                    for col_idx, col_name in [(2, "AM In"), (3, "AM Out"), (4, "PM In"), (5, "PM Out")]:
                        cell = ws.cell(row=current_row, column=col_idx)
                        val = row_data[col_name]
                        cell.value = "" if pd.isna(val) else str(val)
                        cell.alignment = center
                        cell.border = border
                        cell.font = small_font

                for col in [6, 7]:
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.alignment = center
                    cell.border = border
                    cell.font = small_font

                current_row += 1
        
        # Fill remaining rows
        rows_filled = len(second_half)
        for _ in range(rows_filled, 15):
            for col in range(1, 8):
                cell = ws.cell(row=current_row, column=col)
                cell.value = ""
                cell.border = border
                cell.font = small_font
            current_row += 1
        
        # TOTAL row for second DTR
        total_cell = ws.cell(row=current_row, column=1, value="TOTAL")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        total_cell.alignment = center
        total_cell.font = bold
        total_cell.border = border
        
        for col in [6, 7]:
            cell = ws.cell(row=current_row, column=col)
            cell.value = ""
            cell.border = border
        
        current_row += 4
        
        # -------- FOOTER (CERTIFICATION) --------
        footer_cell = ws.cell(row=current_row, column=1, 
                              value="I certify on my honor that the above is a true and correct report of the\n"
                                    "hours of work performed, record of which was made daily at the time of\n"
                                    "arrival and departure from office.")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 2, end_column=7)
        footer_cell.alignment = center
        footer_cell.font = small_font

        current_row += 4
        
        # Signature line
        signature_cell = ws.cell(row=current_row, column=5, value="Principal III")
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
        signature_cell.alignment = center
        signature_cell.font = small_font

        # -------- PAGE SETUP --------
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToHeight = 1
        ws.page_setup.fitToWidth = 1
        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.5
        ws.page_margins.bottom = 0.5
        ws.page_setup.horizontalCentered = True

        # -------- SAVE AND DOWNLOAD --------
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        st.success("✅ DTR Excel file generated successfully!")
        
        safe_name = "".join([c if c.isalnum() or c in "._- " else "_" for c in employee_name])
        
        st.download_button(
            "📥 Download Excel File",
            buffer.getvalue(),
            file_name=f"DTR_CSForm48_{safe_name}_{month}_{year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"❌ Error generating Excel file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
