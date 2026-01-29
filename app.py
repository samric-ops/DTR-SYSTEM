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
        
        # VERY SMALL FONTS to fit everything in one page
        tiny_font = Font(size=7)
        small_font = Font(size=8)
        normal_font = Font(size=9)
        header_font = Font(bold=True, size=8)

        # Set EXTREMELY NARROW column widths for two complete DTRs side by side
        # LEFT DTR: Columns A-G, RIGHT DTR: Columns I-O (H is spacer)
        # We need to make columns very narrow to fit 31 rows + header
        
        column_widths = [2.5, 5, 5, 5, 5, 5, 5]  # Super narrow!
        
        # Set left DTR columns (A-G)
        for i, w in enumerate(column_widths, 1):
            col_letter = chr(64 + i)
            ws.column_dimensions[col_letter].width = w
        
        # Set spacer column H (very narrow spacer)
        ws.column_dimensions['H'].width = 1
        
        # Set right DTR columns (I-O)
        for i, w in enumerate(column_widths, 1):
            col_letter = chr(64 + 8 + i)  # Start from I (9th column)
            ws.column_dimensions[col_letter].width = w
        
        # Set row heights to be smaller
        for row in range(1, 100):
            ws.row_dimensions[row].height = 12  # Smaller row height

        # -------- FUNCTION TO CREATE ONE COMPLETE DTR (1-31 days) --------
        def create_complete_dtr(start_col, start_row):
            """Create one complete DTR with all days (1-31)"""
            r = start_row
            start_col_num = start_col
            
            # HEADER - COMPACT VERSION
            # Line 1-4: Combined to save space
            cell = ws.cell(row=r, column=start_col_num, 
                          value="REPUBLIC OF THE PHILIPPINES\nDepartment of Education\nDivision of Davao del Sur\nManual National High School")
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r, end_column=start_col_num+6)
            cell.alignment = center
            cell.font = Font(bold=True, size=7)
            r += 1
            
            # Civil Service Form and Employee No. - One line
            cell = ws.cell(row=r, column=start_col_num, 
                          value=f"Civil Service Form No. 48{' '*20}Employee No. {employee_no}")
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r, end_column=start_col_num+6)
            cell.alignment = center
            cell.font = Font(bold=True, size=7)
            r += 1
            
            # DAILY TIME RECORD
            cell = ws.cell(row=r, column=start_col_num, value="DAILY TIME RECORD")
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r, end_column=start_col_num+6)
            cell.alignment = center
            cell.font = Font(bold=True, size=9)
            r += 1
            
            # ---o0o--- line
            cell = ws.cell(row=r, column=start_col_num, value="---o0o---")
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r, end_column=start_col_num+6)
            cell.alignment = center
            cell.font = Font(bold=True, size=7)
            r += 1
            
            # Employee Name
            cell = ws.cell(row=r, column=start_col_num, value=employee_name)
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r, end_column=start_col_num+6)
            cell.alignment = center
            cell.font = Font(bold=True, size=8)
            r += 1
            
            # "(Name)" and month/year in one line
            cell = ws.cell(row=r, column=start_col_num, value=f"(Name)   For {month} {year}")
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r, end_column=start_col_num+6)
            cell.alignment = center
            cell.font = Font(size=7)
            r += 1
            
            # Official hours - compact
            cell = ws.cell(row=r, column=start_col_num, 
                          value=f"Hours: {am_hours} / {pm_hours}   Sat: {saturday_hours}")
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r, end_column=start_col_num+6)
            cell.alignment = center
            cell.font = Font(size=7)
            r += 2  # Small space before table
            
            # TABLE HEADER - COMPACT
            # Day header (vertical merge)
            day_cell = ws.cell(row=r, column=start_col_num, value="Day")
            ws.merge_cells(start_row=r, start_column=start_col_num, end_row=r+1, end_column=start_col_num)
            day_cell.alignment = center
            day_cell.font = Font(bold=True, size=7)
            
            # A.M. header (horizontal merge)
            am_cell = ws.cell(row=r, column=start_col_num+1, value="A.M.")
            ws.merge_cells(start_row=r, start_column=start_col_num+1, end_row=r, end_column=start_col_num+2)
            am_cell.alignment = center
            am_cell.font = Font(bold=True, size=7)
            
            # P.M. header (horizontal merge)
            pm_cell = ws.cell(row=r, column=start_col_num+3, value="P.M.")
            ws.merge_cells(start_row=r, start_column=start_col_num+3, end_row=r, end_column=start_col_num+4)
            pm_cell.alignment = center
            pm_cell.font = Font(bold=True, size=7)
            
            # Undertime header (horizontal merge)
            under_cell = ws.cell(row=r, column=start_col_num+5, value="Undertime")
            ws.merge_cells(start_row=r, start_column=start_col_num+5, end_row=r, end_column=start_col_num+6)
            under_cell.alignment = center
            under_cell.font = Font(bold=True, size=7)
            
            # Second row sub-headers
            r += 1
            sub_headers = ["", "In", "Out", "In", "Out", "Hrs", "Min"]  # Shorter labels
            
            for col_idx in range(7):  # 0-6 for 7 columns
                if col_idx != 0:  # Skip column 1 (already has "Day" from merge)
                    cell = ws.cell(row=r, column=start_col_num+col_idx)
                    cell.value = sub_headers[col_idx]
                    cell.alignment = center
                    cell.font = Font(bold=True, size=6)
                    cell.border = border
                else:
                    # Just add border to Day column
                    ws.cell(row=r, column=start_col_num).border = border
            
            return r + 1  # Return next row after header

        # -------- CREATE LEFT DTR HEADER (Columns A-G) --------
        table_start_row_left = create_complete_dtr(1, 1)
        
        # -------- CREATE RIGHT DTR HEADER (Columns I-O) --------
        table_start_row_right = create_complete_dtr(9, 1)
        
        # Use the same starting row for both tables
        table_start_row = max(table_start_row_left, table_start_row_right)
        
        # -------- FILL TABLE DATA FOR LEFT DTR (1-31) --------
        current_row = table_start_row
        
        for day_num in range(1, 32):  # 1 to 31
            if day_num <= len(edited_df):
                row_data = edited_df.iloc[day_num - 1]
                
                # Day number (Column A)
                day_cell = ws.cell(row=current_row, column=1)
                day_val = row_data["Day"]
                day_cell.value = int(day_val) if not pd.isna(day_val) else day_num
                day_cell.alignment = center
                day_cell.border = border
                day_cell.font = tiny_font
                
                # Time entries
                if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                    sat_cell = ws.cell(row=current_row, column=2, value=str(row_data["AM In"]).strip())
                    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
                    sat_cell.alignment = center
                    sat_cell.border = border
                    sat_cell.font = tiny_font
                    
                    # Clear other cells in merge range
                    for col in [3, 4, 5]:
                        ws.cell(row=current_row, column=col).border = border
                else:
                    for col_offset, col_name in [(0, "AM In"), (1, "AM Out"), (2, "PM In"), (3, "PM Out")]:
                        cell = ws.cell(row=current_row, column=2+col_offset)
                        val = row_data[col_name]
                        cell.value = "" if pd.isna(val) else str(val)
                        cell.alignment = center
                        cell.border = border
                        cell.font = tiny_font
                
                # Undertime columns (empty)
                for col in [6, 7]:
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.alignment = center
                    cell.border = border
                    cell.font = tiny_font
            else:
                # Empty row for non-existent days
                for col in range(1, 8):
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.border = border
                    cell.font = tiny_font
            
            current_row += 1
        
        # TOTAL row for left DTR
        total_cell_left = ws.cell(row=current_row, column=1, value="TOTAL")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        total_cell_left.alignment = center
        total_cell_left.font = Font(bold=True, size=7)
        total_cell_left.border = border
        
        for col in [6, 7]:
            cell = ws.cell(row=current_row, column=col)
            cell.value = ""
            cell.border = border
        
        left_dtr_end_row = current_row
        
        # -------- FILL TABLE DATA FOR RIGHT DTR (1-31) --------
        current_row = table_start_row
        
        for day_num in range(1, 32):  # 1 to 31
            if day_num <= len(edited_df):
                row_data = edited_df.iloc[day_num - 1]
                
                # Day number (Column I)
                day_cell = ws.cell(row=current_row, column=9)
                day_val = row_data["Day"]
                day_cell.value = int(day_val) if not pd.isna(day_val) else day_num
                day_cell.alignment = center
                day_cell.border = border
                day_cell.font = tiny_font
                
                # Time entries
                if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                    sat_cell = ws.cell(row=current_row, column=10, value=str(row_data["AM In"]).strip())
                    ws.merge_cells(start_row=current_row, start_column=10, end_row=current_row, end_column=13)
                    sat_cell.alignment = center
                    sat_cell.border = border
                    sat_cell.font = tiny_font
                else:
                    for col_offset, col_name in [(0, "AM In"), (1, "AM Out"), (2, "PM In"), (3, "PM Out")]:
                        cell = ws.cell(row=current_row, column=10+col_offset)
                        val = row_data[col_name]
                        cell.value = "" if pd.isna(val) else str(val)
                        cell.alignment = center
                        cell.border = border
                        cell.font = tiny_font
                
                # Undertime columns (empty)
                for col in [14, 15]:
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.alignment = center
                    cell.border = border
                    cell.font = tiny_font
            else:
                # Empty row for non-existent days
                for col in range(9, 16):
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.border = border
                    cell.font = tiny_font
            
            current_row += 1
        
        # TOTAL row for right DTR
        total_cell_right = ws.cell(row=current_row, column=9, value="TOTAL")
        ws.merge_cells(start_row=current_row, start_column=9, end_row=current_row, end_column=13)
        total_cell_right.alignment = center
        total_cell_right.font = Font(bold=True, size=7)
        total_cell_right.border = border
        
        for col in [14, 15]:
            cell = ws.cell(row=current_row, column=col)
            cell.value = ""
            cell.border = border
        
        right_dtr_end_row = current_row
        
        # Use the maximum end row
        final_table_row = max(left_dtr_end_row, right_dtr_end_row)
        
        # -------- FOOTER (CERTIFICATION) - SPAN BOTH DTRs --------
        footer_start_row = final_table_row + 2
        
        footer_cell = ws.cell(row=footer_start_row, column=1, 
                              value="I certify on my honor that the above is a true and correct report of the "
                                    "hours of work performed, record of which was made daily at the time of "
                                    "arrival and departure from office.")
        # Span from column A to O (across both DTRs)
        ws.merge_cells(start_row=footer_start_row, start_column=1, end_row=footer_start_row, end_column=15)
        footer_cell.alignment = center
        footer_cell.font = Font(size=7)
        
        # -------- SIGNATURE LINE - CENTERED --------
        signature_row = footer_start_row + 2
        signature_cell = ws.cell(row=signature_row, column=7, value="Principal III")
        ws.merge_cells(start_row=signature_row, start_column=7, end_row=signature_row, end_column=9)
        signature_cell.alignment = center
        signature_cell.font = Font(size=7)
        
        final_row = signature_row + 1

        # -------- PAGE SETUP FOR A4 PORTRAIT - ONE PAGE --------
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.fitToHeight = 0  # Don't fit to page height
        ws.page_setup.fitToWidth = 1   # Fit to page width
        ws.page_setup.scale = 85       # Scale down to 85% to fit everything
        
        # Very tight margins to maximize space
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2
        ws.page_margins.header = 0.1
        ws.page_margins.footer = 0.1
        
        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = False
        
        # Force everything to fit on one page
        ws.page_setup.fitToPage = True
        
        # Set print area for the entire page
        ws.print_area = f'A1:O{final_row}'

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
        
        # Show preview note
        st.info("""
        **📝 Printing Instructions:**
        - Paper: **A4 Portrait (One Page Only)**
        - Two **COMPLETE DTRs** (1-31 days each) side by side
        - Left side: Full DTR 1-31
        - Right side: Full DTR 1-31 (duplicate)
        - **Very small fonts** used to fit everything
        - When printing: Use **"Fit to Page"** or **85% scale**
        """)
        
    except Exception as e:
        st.error(f"❌ Error generating Excel file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
