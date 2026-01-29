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
            "REPUBLIC OF THE PHILIPPINES",
            "Department of Education",
            "Division of Davao del Sur",
            "MANUAL NATIONAL HIGH SCHOOL",
            "",  # Empty row
            "DAILY TIME RECORD",
            "-----o0o-----",
            "",  # Empty row
            f"Name: {employee_name}",
            f"For the month of: {month} {year}",
            "",  # Empty row
            "Official hours for arrival and departure",
            f"Regular days: {am_hours} / {pm_hours}",
            f"Saturdays: {saturday_hours}",
            "",  # Empty row
            ""   # Empty row
        ]
        
        for text in header_texts:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
            cell = ws.cell(row=current_row, column=1)
            cell.value = text if text is not None else ""
            cell.alignment = center
            if text and text not in ["", "-----o0o-----"]:
                cell.font = bold
            current_row += 1

        # -------- TABLE HEADER --------
        # Top row headers
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 1, end_column=1)
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
        ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=5)
        ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=7)

        # Main headers
        header_data = [
            (1, "Day"),
            (2, "A.M."),
            (4, "P.M."),
            (6, "Undertime")
        ]
        
        for col, text in header_data:
            cell = ws.cell(row=current_row, column=col)
            cell.value = text
            cell.alignment = center
            cell.font = bold

        current_row += 1
        
        # Sub-headers
        sub_headers = ["", "Arrival", "Departure", "Arrival", "Departure", "Hours", "Minutes"]
        for col_idx in range(1, 8):  # Columns 1-7
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = sub_headers[col_idx - 1] if col_idx - 1 < len(sub_headers) else ""
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
            if row_data["AM In"] in ["SATURDAY", "SUNDAY"]:
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
                merged_cell = ws.cell(row=current_row, column=2)
                merged_cell.value = str(row_data["AM In"])
                merged_cell.alignment = center
                merged_cell.border = border
            else:
                # AM In
                am_in_cell = ws.cell(row=current_row, column=2)
                am_in_cell.value = "" if pd.isna(row_data["AM In"]) else str(row_data["AM In"])
                am_in_cell.alignment = center
                am_in_cell.border = border
                
                # AM Out
                am_out_cell = ws.cell(row=current_row, column=3)
                am_out_cell.value = "" if pd.isna(row_data["AM Out"]) else str(row_data["AM Out"])
                am_out_cell.alignment = center
                am_out_cell.border = border
                
                # PM In
                pm_in_cell = ws.cell(row=current_row, column=4)
                pm_in_cell.value = "" if pd.isna(row_data["PM In"]) else str(row_data["PM In"])
                pm_in_cell.alignment = center
                pm_in_cell.border = border
                
                # PM Out
                pm_out_cell = ws.cell(row=current_row, column=5)
                pm_out_cell.value = "" if pd.isna(row_data["PM Out"]) else str(row_data["PM Out"])
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
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        total_cell = ws.cell(row=current_row, column=1)
        total_cell.value = "TOTAL"
        total_cell.alignment = center
        total_cell.font = bold
        
        # Add border to all cells in total row
        for col in range(1, 8):
            cell = ws.cell(row=current_row, column=col)
            cell.border = border

        current_row += 3

        # -------- FOOTER --------
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row + 2, end_column=7)
        footer_cell = ws.cell(row=current_row, column=1)
        footer_cell.value = (
            "I certify on my honor that the above is a true and correct report of the\n"
            "hours of work performed, record of which was made daily at the time of\n"
            "arrival and departure from office."
        )
        footer_cell.alignment = center

        current_row += 4
        
        # Signature line
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7)
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
