import streamlit as st
import pandas as pd
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
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
        
        # Remove default gridlines
        ws.sheet_view.showGridLines = False

        # Define styles - USING ARIAL FONT
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left = Alignment(horizontal="left", vertical="center")
        right = Alignment(horizontal="right", vertical="center")
        
        # Arial Font definitions
        font_arial_6 = Font(name='Arial', size=6)
        font_arial_7 = Font(name='Arial', size=7)
        font_arial_8 = Font(name='Arial', size=8)
        font_arial_8_bold = Font(name='Arial', size=8, bold=True)
        font_arial_9 = Font(name='Arial', size=9)
        font_arial_10 = Font(name='Arial', size=10)
        font_arial_10_bold = Font(name='Arial', size=10, bold=True)
        font_arial_10_bold_underline = Font(name='Arial', size=10, bold=True, underline='single')
        font_arial_11_bold = Font(name='Arial', size=11, bold=True)
        font_arial_12_bold = Font(name='Arial', size=12, bold=True)
        font_arial_8_italic = Font(name='Arial', size=8, italic=True)
        
        thin = Side(style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Set column widths for TWO DTRs side by side
        # LEFT DTR: Columns A-J (10 columns for left DTR)
        # RIGHT DTR: Columns L-U (10 columns for right DTR, K is spacer)
        
        # Column widths based on your specification
        dtr_columns = [
            2.5,   # A/L: Day (narrow)
            3.0,   # B/M: A.M. header
            3.0,   # C/N: Arrival
            3.0,   # D/O: Departure
            3.0,   # E/P: P.M. header
            3.0,   # F/Q: Arrival
            3.0,   # G/R: Departure
            3.5,   # H/S: Undertime header
            2.5,   # I/T: Hours
            2.5    # J/U: Minutes
        ]
        
        # Set left DTR columns (A-J)
        for i, width in enumerate(dtr_columns, 1):
            col_letter = get_column_letter(i)
            ws.column_dimensions[col_letter].width = width
        
        # Set spacer column K
        ws.column_dimensions['K'].width = 1.0
        
        # Set right DTR columns (L-U)
        for i, width in enumerate(dtr_columns, 1):
            col_letter = get_column_letter(11 + i)  # L is 12th column (11+1)
            ws.column_dimensions[col_letter].width = width
        
        # Set row heights
        ws.row_dimensions[1].height = 12
        ws.row_dimensions[2].height = 12
        ws.row_dimensions[3].height = 12
        ws.row_dimensions[4].height = 12
        ws.row_dimensions[5].height = 6  # Small space
        ws.row_dimensions[6].height = 12
        ws.row_dimensions[7].height = 12
        ws.row_dimensions[8].height = 12
        ws.row_dimensions[9].height = 12
        ws.row_dimensions[10].height = 12
        ws.row_dimensions[11].height = 12
        ws.row_dimensions[12].height = 12
        ws.row_dimensions[13].height = 12
        ws.row_dimensions[14].height = 15  # Table header
        
        # Set table rows to 12
        for row in range(15, 50):
            ws.row_dimensions[row].height = 12

        # -------- FUNCTION TO CREATE ONE DTR --------
        def create_dtr(start_col, start_row):
            """Create one complete DTR starting at specified column"""
            r = start_row
            col = start_col
            
            # 1. REPUBLIC OF THE PHILIPPINES Header (4 lines combined)
            cell = ws.cell(row=r, column=col, 
                          value="REPUBLIC OF THE PHILIPPINES\nDepartment of Education\nDivision of Davao del Sur\nManual National High School")
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+9)
            cell.alignment = center
            cell.font = font_arial_8
            r += 1
            
            # 2. Blank line (0.3 space)
            r += 1
            
            # 3. Civil Service Form No. 48 and Employee No.
            # Left part
            cell_left = ws.cell(row=r, column=col, value="Civil Service Form No. 48")
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+4)
            cell_left.alignment = left
            cell_left.font = font_arial_8
            
            # Right part with underline for Employee No.
            cell_right = ws.cell(row=r, column=col+5, value=f"Employee No. {employee_no}")
            ws.merge_cells(start_row=r, start_column=col+5, end_row=r, end_column=col+9)
            cell_right.alignment = right
            cell_right.font = font_arial_8
            # Add underline (simulated with border bottom)
            for c in range(col+5, col+10):
                ws.cell(row=r, column=c).border = Border(bottom=thin)
            r += 1
            
            # 4. Blank line (0.3 space)
            r += 1
            
            # 5. DAILY TIME RECORD (centered, size 12, bold)
            cell = ws.cell(row=r, column=col, value="DAILY TIME RECORD")
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+9)
            cell.alignment = center
            cell.font = font_arial_12_bold
            r += 1
            
            # 6. -----o0o----- (no space downward)
            cell = ws.cell(row=r, column=col, value="-----o0o-----")
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+9)
            cell.alignment = center
            cell.font = font_arial_8
            r += 1
            
            # 7. Employee Name (single line space, centered, size 10-11, bold, underlined)
            cell = ws.cell(row=r, column=col, value=employee_name)
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+9)
            cell.alignment = center
            cell.font = font_arial_10_bold_underline
            r += 1
            
            # 8. "(Name)" label (centered, size 7-8, not bold)
            cell = ws.cell(row=r, column=col, value="(Name)")
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+9)
            cell.alignment = center
            cell.font = font_arial_7
            r += 1
            
            # 9. Period and Official Hours (split into two parts)
            # Left part: For the month of
            cell_left = ws.cell(row=r, column=col, value=f"For the month of {month.upper()} {year}")
            ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+4)
            cell_left.alignment = left
            cell_left.font = font_arial_8
            
            # Right part: Official hours
            hours_text = f"Official hours for arrival and departure\nRegular days: {am_hours} / {pm_hours}\nSaturdays: {saturday_hours}"
            cell_right = ws.cell(row=r, column=col+5, value=hours_text)
            ws.merge_cells(start_row=r, start_column=col+5, end_row=r+2, end_column=col+9)
            cell_right.alignment = left
            cell_right.font = font_arial_8
            r += 3  # Move down 3 rows for the multi-line hours text
            
            # 10. Blank line before table
            r += 1
            
            # Store table start row
            table_start_row = r
            
            # 11. TABLE HEADER
            # Day column (merged vertically for 2 rows)
            day_cell = ws.cell(row=r, column=col, value="Day")
            ws.merge_cells(start_row=r, start_column=col, end_row=r+1, end_column=col)
            day_cell.alignment = center
            day_cell.font = font_arial_8_bold
            day_cell.border = border
            
            # A.M. header (merged horizontally for 2 columns)
            am_cell = ws.cell(row=r, column=col+1, value="A.M.")
            ws.merge_cells(start_row=r, start_column=col+1, end_row=r, end_column=col+2)
            am_cell.alignment = center
            am_cell.font = font_arial_8_bold
            am_cell.border = border
            
            # P.M. header (merged horizontally for 2 columns)
            pm_cell = ws.cell(row=r, column=col+4, value="P.M.")
            ws.merge_cells(start_row=r, start_column=col+4, end_row=r, end_column=col+5)
            pm_cell.alignment = center
            pm_cell.font = font_arial_8_bold
            pm_cell.border = border
            
            # Undertime header (merged horizontally for 2 columns)
            under_cell = ws.cell(row=r, column=col+7, value="Undertime")
            ws.merge_cells(start_row=r, start_column=col+7, end_row=r, end_column=col+8)
            under_cell.alignment = center
            under_cell.font = font_arial_8_bold
            under_cell.border = border
            
            # Second row of headers
            r += 1
            
            # Sub-headers for A.M.
            arrival_am = ws.cell(row=r, column=col+1, value="Arrival")
            arrival_am.alignment = center
            arrival_am.font = font_arial_8_bold
            arrival_am.border = border
            
            departure_am = ws.cell(row=r, column=col+2, value="Departure")
            departure_am.alignment = center
            departure_am.font = font_arial_8_bold
            departure_am.border = border
            
            # Sub-headers for P.M.
            arrival_pm = ws.cell(row=r, column=col+4, value="Arrival")
            arrival_pm.alignment = center
            arrival_pm.font = font_arial_8_bold
            arrival_pm.border = border
            
            departure_pm = ws.cell(row=r, column=col+5, value="Departure")
            departure_pm.alignment = center
            departure_pm.font = font_arial_8_bold
            departure_pm.border = border
            
            # Sub-headers for Undertime
            hours = ws.cell(row=r, column=col+7, value="Hours")
            hours.alignment = center
            hours.font = font_arial_8_bold
            hours.border = border
            
            minutes = ws.cell(row=r, column=col+8, value="Minutes")
            minutes.alignment = center
            minutes.font = font_arial_8_bold
            minutes.border = border
            
            # Add borders to empty cells in header
            ws.cell(row=r, column=col).border = border  # Day column bottom
            ws.cell(row=r, column=col+3).border = border  # Spacer between AM/PM
            ws.cell(row=r, column=col+6).border = border  # Spacer between PM/Undertime
            ws.cell(row=r, column=col+9).border = border  # Empty cell
            
            return table_start_row + 2  # Return first data row

        # -------- CREATE LEFT DTR (Columns A-J) --------
        data_start_left = create_dtr(1, 1)
        
        # -------- CREATE RIGHT DTR (Columns L-U) --------
        data_start_right = create_dtr(12, 1)  # Column L is 12
        
        # Make sure both start at same row
        data_start_row = max(data_start_left, data_start_right)
        
        # -------- FILL TABLE DATA FOR BOTH DTRs --------
        # We'll fill 31 rows for both DTRs
        current_row = data_start_row
        
        for day_num in range(1, 32):  # 1 to 31
            if day_num <= len(edited_df):
                row_data = edited_df.iloc[day_num - 1]
                
                # ===== LEFT DTR =====
                # Day number
                day_cell_left = ws.cell(row=current_row, column=1, value=day_num)
                day_cell_left.alignment = center
                day_cell_left.font = font_arial_8
                day_cell_left.border = border
                
                # Time entries for LEFT DTR
                if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                    # Merge for SATURDAY/SUNDAY
                    sat_cell = ws.cell(row=current_row, column=2, value=str(row_data["AM In"]).strip())
                    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
                    sat_cell.alignment = center
                    sat_cell.font = font_arial_8
                    sat_cell.border = border
                    
                    # Add borders to merged cells
                    for col in [3, 4, 5]:
                        ws.cell(row=current_row, column=col).border = border
                else:
                    # Regular day - fill all time cells
                    for col_offset, col_name in [(0, "AM In"), (1, "AM Out"), (2, "PM In"), (3, "PM Out")]:
                        col_idx = 2 + col_offset
                        cell = ws.cell(row=current_row, column=col_idx)
                        val = row_data[col_name]
                        cell.value = "" if pd.isna(val) else str(val)
                        cell.alignment = center
                        cell.font = font_arial_8
                        cell.border = border
                
                # Undertime columns for LEFT DTR (empty)
                for col in [7, 8]:  # Hours and Minutes columns
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.alignment = center
                    cell.font = font_arial_8
                    cell.border = border
                
                # Empty cells
                ws.cell(row=current_row, column=6).border = border  # Spacer
                ws.cell(row=current_row, column=9).border = border  # Empty
                
                # ===== RIGHT DTR =====
                # Day number
                day_cell_right = ws.cell(row=current_row, column=12, value=day_num)
                day_cell_right.alignment = center
                day_cell_right.font = font_arial_8
                day_cell_right.border = border
                
                # Time entries for RIGHT DTR
                if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                    # Merge for SATURDAY/SUNDAY
                    sat_cell = ws.cell(row=current_row, column=13, value=str(row_data["AM In"]).strip())
                    ws.merge_cells(start_row=current_row, start_column=13, end_row=current_row, end_column=16)
                    sat_cell.alignment = center
                    sat_cell.font = font_arial_8
                    sat_cell.border = border
                    
                    # Add borders to merged cells
                    for col in [14, 15, 16]:
                        ws.cell(row=current_row, column=col).border = border
                else:
                    # Regular day - fill all time cells
                    for col_offset, col_name in [(0, "AM In"), (1, "AM Out"), (2, "PM In"), (3, "PM Out")]:
                        col_idx = 13 + col_offset
                        cell = ws.cell(row=current_row, column=col_idx)
                        val = row_data[col_name]
                        cell.value = "" if pd.isna(val) else str(val)
                        cell.alignment = center
                        cell.font = font_arial_8
                        cell.border = border
                
                # Undertime columns for RIGHT DTR (empty)
                for col in [18, 19]:  # Hours and Minutes columns
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.alignment = center
                    cell.font = font_arial_8
                    cell.border = border
                
                # Empty cells
                ws.cell(row=current_row, column=17).border = border  # Spacer
                ws.cell(row=current_row, column=20).border = border  # Empty
                
            else:
                # Empty row if day doesn't exist (fill with borders)
                for col in range(1, 11):  # Left DTR columns
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.border = border
                
                for col in range(12, 22):  # Right DTR columns
                    cell = ws.cell(row=current_row, column=col)
                    cell.value = ""
                    cell.border = border
            
            current_row += 1
        
        # ===== TOTAL ROW FOR BOTH DTRs =====
        # LEFT DTR TOTAL
        total_left = ws.cell(row=current_row, column=1, value="TOTAL")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        total_left.alignment = center
        total_left.font = font_arial_8_bold
        total_left.border = border
        
        # Add borders to remaining cells in LEFT DTR
        for col in range(6, 11):
            cell = ws.cell(row=current_row, column=col)
            cell.value = ""
            cell.border = border
        
        # RIGHT DTR TOTAL
        total_right = ws.cell(row=current_row, column=12, value="TOTAL")
        ws.merge_cells(start_row=current_row, start_column=12, end_row=current_row, end_column=16)
        total_right.alignment = center
        total_right.font = font_arial_8_bold
        total_right.border = border
        
        # Add borders to remaining cells in RIGHT DTR
        for col in range(17, 22):
            cell = ws.cell(row=current_row, column=col)
            cell.value = ""
            cell.border = border
        
        # ===== FOOTER SECTION =====
        footer_start = current_row + 2
        
        # Certification text (spans both DTRs)
        cert_cell = ws.cell(row=footer_start, column=1, 
                           value="I certify on my honor that the above is a true and correct report of the hours of work performed, record of which was made daily at the time of arrival and departure from office.")
        ws.merge_cells(start_row=footer_start, start_column=1, end_row=footer_start, end_column=21)
        cert_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cert_cell.font = font_arial_8_italic
        
        # Signature line (2 spaces down)
        signature_line_row = footer_start + 4
        
        # Horizontal line for signature (centered, spans both DTRs)
        sig_line = ws.cell(row=signature_line_row, column=8, value="_________________________")
        ws.merge_cells(start_row=signature_line_row, start_column=8, end_row=signature_line_row, end_column=14)
        sig_line.alignment = center
        sig_line.font = font_arial_8
        
        # "VERIFIED as to the prescribed office hours:" (1 space down)
        verified_row = signature_line_row + 2
        verified_cell = ws.cell(row=verified_row, column=1, value="VERIFIED as to the prescribed office hours:")
        ws.merge_cells(start_row=verified_row, start_column=1, end_row=verified_row, end_column=21)
        verified_cell.alignment = center
        verified_cell.font = font_arial_8
        
        # Space for principal signature
        principal_sig_row = verified_row + 3
        
        # Principal III (centered, should be near bottom margin)
        principal_cell = ws.cell(row=principal_sig_row, column=8, value="Principal III")
        ws.merge_cells(start_row=principal_sig_row, start_column=8, end_row=principal_sig_row, end_column=14)
        principal_cell.alignment = center
        principal_cell.font = font_arial_8_bold
        
        # Add horizontal line above Principal III
        for col in range(8, 15):
            ws.cell(row=principal_sig_row-1, column=col).border = Border(top=thin)
        
        final_row = principal_sig_row + 1

        # ===== PAGE SETUP =====
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        
        # Set margins as per specification (convert cm to inches)
        # 1 cm = 0.393701 inches
        ws.page_margins.left = 1.27 * 0.393701  # 1.27 cm to inches
        ws.page_margins.right = 1.27 * 0.393701  # 1.27 cm to inches
        ws.page_margins.top = 1.9 * 0.393701    # 1.9 cm to inches
        ws.page_margins.bottom = 0.49 * 0.393701  # 0.49 cm to inches
        ws.page_margins.header = 0.3
        ws.page_margins.footer = 0.3
        
        ws.page_setup.horizontalCentered = True
        ws.page_setup.verticalCentered = False
        
        # Set print area to ensure everything fits
        ws.print_area = f'A1:U{final_row}'
        
        # Scale to fit (if needed)
        ws.page_setup.fitToHeight = 0
        ws.page_setup.fitToWidth = 1
        ws.page_setup.scale = 100  # 100% scale

        # ===== SAVE AND DOWNLOAD =====
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
        **📝 Document Specifications:**
        - Paper: **A4 Portrait**
        - Margins: Top 1.9cm, Bottom 0.49cm, Left/Right 1.27cm
        - Font: **Arial** throughout
        - Two identical DTRs side by side
        - Each DTR contains complete 1-31 days
        - Proper formatting as per Civil Service Form 48
        """)
        
    except Exception as e:
        st.error(f"❌ Error generating Excel file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
