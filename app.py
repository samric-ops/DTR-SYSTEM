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

    wb = Workbook()
    ws = wb.active
    ws.title = "DTR"

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bold = Font(bold=True)
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    widths = [6, 10, 10, 10, 10, 10, 10]
    for i, w in enumerate(widths, 1):
        col_letter = chr(64 + i) if i <= 26 else chr(64 + (i // 26)) + chr(64 + (i % 26))
        ws.column_dimensions[col_letter].width = w

    r = 1

    # ---------- REVISED write_merged FUNCTION ----------
    def write_merged(text, rows=1, is_bold=True):
        nonlocal r
        start_row = r
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row + rows - 1, end_column=7)
        cell = ws.cell(row=start_row, column=1)
        cell.value = "" if text is None else str(text)
        cell.alignment = center
        if is_bold:
            cell.font = bold
        r += rows

    # -------- HEADER --------
    write_merged("REPUBLIC OF THE PHILIPPINES")
    write_merged("Department of Education")
    write_merged("Division of Davao del Sur")
    write_merged("MANUAL NATIONAL HIGH SCHOOL")
    write_merged("", rows=1, is_bold=False)  # Empty row
    write_merged("DAILY TIME RECORD")
    write_merged("-----o0o-----")
    write_merged("", rows=1, is_bold=False)  # Empty row
    write_merged(f"Name: {employee_name}")
    write_merged(f"For the month of: {month} {year}")
    write_merged("", rows=1, is_bold=False)  # Empty row
    write_merged("Official hours for arrival and departure")
    write_merged(f"Regular days: {am_hours} / {pm_hours}")
    write_merged(f"Saturdays: {saturday_hours}")
    write_merged("", rows=2, is_bold=False)  # 2 empty rows

    # -------- TABLE HEADER --------
    # Top row headers
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=1)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)

    headers = {1: "Day", 2: "A.M.", 4: "P.M.", 6: "Undertime"}
    for col, text in headers.items():
        cell = ws.cell(row=r, column=col)
        cell.value = text
        cell.alignment = center
        cell.font = bold

    # Second row sub-headers
    r += 1
    sub_headers = ["", "Arrival", "Departure", "Arrival", "Departure", "Hours", "Minutes"]
    
    # I-debug muna kung may laman ang sub_headers
    if not sub_headers or len(sub_headers) == 0:
        st.error("Error: sub_headers is empty!")
        st.stop()
    
    for c, text in enumerate(sub_headers, start=1):
        cell = ws.cell(row=r, column=c)
        cell.value = text
        cell.alignment = center
        cell.font = Font(bold=True)
        cell.border = border

    r += 1

    # -------- TABLE DATA --------
    for _, row in edited_df.iterrows():
        # Day column
        cell_day = ws.cell(row=r, column=1)
        cell_day.value = row["Day"]
        cell_day.alignment = center
        cell_day.border = border

        if row["AM In"] in ["SATURDAY", "SUNDAY"]:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            cell_merged = ws.cell(row=r, column=2)
            cell_merged.value = row["AM In"]
            cell_merged.alignment = center
            cell_merged.border = border
        else:
            # AM In (Arrival)
            cell_am_in = ws.cell(row=r, column=2)
            cell_am_in.value = "" if pd.isna(row["AM In"]) else str(row["AM In"])
            cell_am_in.alignment = center
            cell_am_in.border = border
            
            # AM Out (Departure)
            cell_am_out = ws.cell(row=r, column=3)
            cell_am_out.value = "" if pd.isna(row["AM Out"]) else str(row["AM Out"])
            cell_am_out.alignment = center
            cell_am_out.border = border
            
            # PM In (Arrival)
            cell_pm_in = ws.cell(row=r, column=4)
            cell_pm_in.value = "" if pd.isna(row["PM In"]) else str(row["PM In"])
            cell_pm_in.alignment = center
            cell_pm_in.border = border
            
            # PM Out (Departure)
            cell_pm_out = ws.cell(row=r, column=5)
            cell_pm_out.value = "" if pd.isna(row["PM Out"]) else str(row["PM Out"])
            cell_pm_out.alignment = center
            cell_pm_out.border = border

        # Undertime columns (blank by default)
        for c in [6, 7]:
            cell = ws.cell(row=r, column=c)
            cell.alignment = center
            cell.border = border

        r += 1

    # -------- TOTAL ROW --------
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    total_cell = ws.cell(row=r, column=1)
    total_cell.value = "TOTAL"
    total_cell.alignment = center
    total_cell.font = bold
    
    # Add border to total row
    for c in range(1, 8):
        ws.cell(row=r, column=c).border = border

    r += 3

    # -------- FOOTER --------
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=7)
    footer_cell = ws.cell(row=r, column=1)
    footer_cell.value = (
        "I certify on my honor that the above is a true and correct report of the\n"
        "hours of work performed, record of which was made daily at the time of\n"
        "arrival and departure from office."
    )
    footer_cell.alignment = center

    r += 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    signature_cell = ws.cell(row=r, column=5)
    signature_cell.value = "Principal III"
    signature_cell.alignment = center

    # -------- PRINT SETTINGS (A4) --------
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToHeight = 1
    ws.page_setup.fitToWidth = 1
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75
    ws.page_setup.horizontalCentered = True
    ws.print_title_rows = "1:18"

    # -------- DOWNLOAD --------
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success("✅ DTR Excel file generated successfully!")
    st.download_button(
        "📥 Download Excel File",
        buffer.getvalue(),
        file_name=f"DTR_{employee_name}_{month}_{year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
