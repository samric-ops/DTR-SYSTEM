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
        ws.column_dimensions[chr(64 + i)].width = w

    r = 1

    def write_merged(text, rows=1, is_bold=True):
        global r
        ws.merge_cells(start_row=r, start_column=1, end_row=r + rows - 1, end_column=7)
        cell = ws.cell(r, 1)
        cell.value = text
        cell.alignment = center
        if is_bold:
            cell.font = bold
        r += rows

    # -------- HEADER --------
    write_merged("REPUBLIC OF THE PHILIPPINES")
    write_merged("Department of Education")
    write_merged("Division of Davao del Sur")
    write_merged("MANUAL NATIONAL HIGH SCHOOL")
    r += 1
    write_merged("DAILY TIME RECORD")
    write_merged("-----o0o-----")
    r += 1
    write_merged(f"Name: {employee_name}")
    write_merged(f"For the month of: {month} {year}")
    r += 1
    write_merged("Official hours for arrival and departure")
    write_merged(f"Regular days: {am_hours} / {pm_hours}")
    write_merged(f"Saturdays: {saturday_hours}")
    r += 2

    # -------- TABLE HEADER --------
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=1)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)

    ws.cell(r, 1).value = "Day"
    ws.cell(r, 2).value = "A.M."
    ws.cell(r, 4).value = "P.M."
    ws.cell(r, 6).value = "Undertime"

    for col in [1, 2, 4, 6]:
        ws.cell(r, col).alignment = center
        ws.cell(r, col).font = bold

    r += 1
    sub_headers = ["", "Arrival", "Departure", "Arrival", "Departure", "Hours", "Minutes"]
    for c, text in enumerate(sub_headers, 1):
        cell = ws.cell(r, c)
        cell.value = text
        cell.alignment = center
        cell.font = bold
        cell.border = border

    r += 1

    # -------- TABLE DATA --------
    for _, row in edited_df.iterrows():
        ws.cell(r, 1).value = row["Day"]
        ws.cell(r, 1).alignment = center

        if row["AM In"] in ["SATURDAY", "SUNDAY"]:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            ws.cell(r, 2).value = row["AM In"]
            ws.cell(r, 2).alignment = center
        else:
            ws.cell(r, 2).value = row["AM In"]
            ws.cell(r, 3).value = row["AM Out"]
            ws.cell(r, 4).value = row["PM In"]
            ws.cell(r, 5).value = row["PM Out"]

        for c in range(1, 8):
            ws.cell(r, c).alignment = center
            ws.cell(r, c).border = border

        r += 1

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.cell(r, 1).value = "TOTAL"
    ws.cell(r, 1).alignment = center
    ws.cell(r, 1).font = bold

    r += 3

    # -------- FOOTER --------
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=7)
    ws.cell(r, 1).value = (
        "I certify on my honor that the above is a true and correct report of the\n"
        "hours of work performed, record of which was made daily at the time of\n"
        "arrival and departure from office."
    )
    ws.cell(r, 1).alignment = center

    r += 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    ws.cell(r, 5).value = "Principal III"
    ws.cell(r, 5).alignment = center

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

    st.success("✅ DTR Excel file generated successfully!")
    st.download_button(
        "📥 Download Excel File",
        buffer.getvalue(),
        file_name=f"DTR_{employee_name}_{month}_{year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
