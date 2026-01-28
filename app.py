import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, Protection
from openpyxl.utils import get_column_letter
import io
import base64
import os
from pathlib import Path
import zipfile

# =============================================
# CIVIL SERVICE FORM NO. 48 GENERATOR FUNCTION
# =============================================

def generate_civil_service_dtr(employee_no, employee_name, month, year, attendance_data, office_hours):
    """Generate Civil Service Form No. 48 in Excel format"""
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f"DTR {employee_name}"
    
    # Styles
    title_font = Font(name='Arial', size=14, bold=True)
    header_font = Font(name='Arial', size=11, bold=True)
    normal_font = Font(name='Arial', size=10)
    small_font = Font(name='Arial', size=9)
    
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # ========== HEADER SECTION ==========
    # Republic of the Philippines
    ws.merge_cells('A1:G1')
    ws['A1'] = "REPUBLIC OF THE PHILIPPINES"
    ws['A1'].font = header_font
    ws['A1'].alignment = center_align
    
    # Department of Education
    ws.merge_cells('A2:G2')
    ws['A2'] = "Department of Education"
    ws['A2'].font = header_font
    ws['A2'].alignment = center_align
    
    # Division
    ws.merge_cells('A3:G3')
    ws['A3'] = "Division of Davao del Sur"
    ws['A3'].font = header_font
    ws['A3'].alignment = center_align
    
    # School
    ws.merge_cells('A4:G4')
    ws['A4'] = "Manual National High School"
    ws['A4'].font = header_font
    ws['A4'].alignment = center_align
    
    # Civil Service Form No. 48
    ws.merge_cells('A6:G6')
    ws['A6'] = "Civil Service Form No. 48"
    ws['A6'].font = small_font
    ws['A6'].alignment = center_align
    
    # Employee No.
    ws.merge_cells('A7:G7')
    ws['A7'] = f"Employee No. {employee_no}"
    ws['A7'].font = small_font
    ws['A7'].alignment = center_align
    
    # DAILY TIME RECORD Title
    ws.merge_cells('A9:G9')
    ws['A9'] = "DAILY TIME RECORD"
    ws['A9'].font = title_font
    ws['A9'].alignment = center_align
    
    # Separator line
    ws.merge_cells('A10:G10')
    ws['A10'] = "-------------------------o0o-------------------------"
    ws['A10'].alignment = center_align
    
    # Employee Name
    ws.merge_cells('A12:G12')
    ws['A12'] = employee_name.upper()
    ws['A12'].font = Font(name='Arial', size=12, bold=True)
    ws['A12'].alignment = center_align
    
    # (Name) label
    ws.merge_cells('A13:G13')
    ws['A13'] = "(Name)"
    ws['A13'].font = small_font
    ws['A13'].alignment = center_align
    
    # ========== MONTH AND OFFICE HOURS SECTION ==========
    current_row = 15
    
    # Month and Year
    month_name = calendar.month_name[month]
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = f"{month_name.upper()} {year}"
    ws[f'A{current_row}'].font = Font(name='Arial', size=11, bold=True)
    ws[f'A{current_row}'].alignment = center_align
    
    ws.merge_cells(f'E{current_row}:G{current_row}')
    ws[f'E{current_row}'] = f"For the month of"
    ws[f'E{current_row}'].font = small_font
    ws[f'E{current_row}'].alignment = left_align
    
    current_row += 1
    
    # Office Hours
    am_hours = f"{office_hours['regular_am_in']} -- {office_hours['regular_am_out']}"
    pm_hours = f"{office_hours['regular_pm_in']} -- {office_hours['regular_pm_out']}"
    
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = am_hours
    ws[f'A{current_row}'].font = normal_font
    ws[f'A{current_row}'].alignment = center_align
    
    ws.merge_cells(f'E{current_row}:F{current_row}')
    ws[f'E{current_row}'] = "Official hours for arrival and departure"
    ws[f'E{current_row}'].font = small_font
    ws[f'E{current_row}'].alignment = left_align
    
    ws[f'G{current_row}'] = "Regular days"
    ws[f'G{current_row}'].font = small_font
    ws[f'G{current_row}'].alignment = left_align
    
    current_row += 1
    
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = pm_hours
    ws[f'A{current_row}'].font = normal_font
    ws[f'A{current_row}'].alignment = center_align
    
    ws.merge_cells(f'E{current_row}:F{current_row}')
    ws[f'E{current_row}'] = ""
    
    ws[f'G{current_row}'] = "Saturdays"
    ws[f'G{current_row}'].font = small_font
    ws[f'G{current_row}'].alignment = left_align
    
    current_row += 1
    
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = office_hours['saturday']
    ws[f'A{current_row}'].font = normal_font
    ws[f'A{current_row}'].alignment = center_align
    
    current_row += 2  # Add spacing
    
    # ========== DTR TABLE HEADER ==========
    # Table header row
    ws.merge_cells(f'A{current_row}:A{current_row+1}')  # Day
    ws[f'A{current_row}'] = "Day"
    ws[f'A{current_row}'].font = header_font
    ws[f'A{current_row}'].alignment = center_align
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:C{current_row}')  # A.M.
    ws[f'B{current_row}'] = "A.M."
    ws[f'B{current_row}'].font = header_font
    ws[f'B{current_row}'].alignment = center_align
    ws[f'B{current_row}'].border = thin_border
    
    ws.merge_cells(f'D{current_row}:E{current_row}')  # P.M.
    ws[f'D{current_row}'] = "P.M."
    ws[f'D{current_row}'].font = header_font
    ws[f'D{current_row}'].alignment = center_align
    ws[f'D{current_row}'].border = thin_border
    
    ws.merge_cells(f'F{current_row}:G{current_row}')  # Undertime
    ws[f'F{current_row}'] = "Undertime"
    ws[f'F{current_row}'].font = header_font
    ws[f'F{current_row}'].alignment = center_align
    ws[f'F{current_row}'].border = thin_border
    
    current_row += 1
    
    # Subheaders
    subheader_cols = ['B', 'C', 'D', 'E', 'F', 'G']
    subheader_texts = ['Arrival', 'Departure', 'Arrival', 'Departure', 'Hours', 'Minutes']
    
    for col, text in zip(subheader_cols, subheader_texts):
        ws[f'{col}{current_row}'] = text
        ws[f'{col}{current_row}'].font = small_font
        ws[f'{col}{current_row}'].alignment = center_align
        ws[f'{col}{current_row}'].border = thin_border
    
    current_row += 1
    
    # ========== DTR TABLE DATA ==========
    # Get days in month
    days_in_month = calendar.monthrange(year, month)[1]
    
    # Process attendance data by day - WITH ERROR HANDLING
    attendance_by_day = {}
    
    # FIXED: Check if attendance_data is valid BEFORE iterrows()
    if attendance_data is not None and not attendance_data.empty and hasattr(attendance_data, 'iterrows'):
        try:
            # FIXED: Make sure we're working with a DataFrame
            if isinstance(attendance_data, pd.DataFrame):
                # FIXED: Reset index to avoid issues
                df_temp = attendance_data.reset_index(drop=True)
                
                # Check for required columns
                if 'Day' in df_temp.columns and 'Time' in df_temp.columns:
                    for index, row in df_temp.iterrows():
                        try:
                            # Get day
                            if pd.notna(row['Day']):
                                day = int(row['Day'])
                            else:
                                continue
                            
                            # Get time
                            if pd.notna(row['Time']):
                                time_val = row['Time']
                                if hasattr(time_val, 'strftime'):
                                    time_str = time_val.strftime('%H:%M')
                                elif isinstance(time_val, str):
                                    time_str = time_val
                                else:
                                    time_str = str(time_val)
                                
                                if day > 0 and day <= 31:
                                    if day not in attendance_by_day:
                                        attendance_by_day[day] = []
                                    attendance_by_day[day].append(time_str)
                        except:
                            continue
        except:
            # If error, continue with empty data
            pass
    
    total_undertime_hours = 0
    total_undertime_minutes = 0
    
    # Fill table for each day
    for day in range(1, days_in_month + 1):
        # Get day of week
        date_obj = datetime(year, month, day)
        day_name = date_obj.strftime('%A')
        
        # Day cell
        ws[f'A{current_row}'] = str(day)
        ws[f'A{current_row}'].font = Font(name='Arial', size=10, bold=True)
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        # Check if Saturday or Sunday
        if day_name.upper() == 'SATURDAY':
            ws.merge_cells(f'B{current_row}:C{current_row}')
            ws[f'B{current_row}'] = "SATURDAY"
            ws[f'B{current_row}'].font = Font(name='Arial', size=10, bold=True, italic=True)
            ws[f'B{current_row}'].alignment = center_align
            ws[f'B{current_row}'].border = thin_border
            ws[f'B{current_row}'].fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
            
            # Clear other cells
            for col in ['D', 'E', 'F', 'G']:
                ws[f'{col}{current_row}'].border = thin_border
        
        elif day_name.upper() == 'SUNDAY':
            ws.merge_cells(f'B{current_row}:C{current_row}')
            ws[f'B{current_row}'] = "SUNDAY"
            ws[f'B{current_row}'].font = Font(name='Arial', size=10, bold=True, italic=True)
            ws[f'B{current_row}'].alignment = center_align
            ws[f'B{current_row}'].border = thin_border
            ws[f'B{current_row}'].fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
            
            # Clear other cells
            for col in ['D', 'E', 'F', 'G']:
                ws[f'{col}{current_row}'].border = thin_border
        
        else:
            # Regular work day - fill with attendance data
            if day in attendance_by_day:
                times = sorted(attendance_by_day[day])
                
                # AM Arrival (first time before 12:00)
                am_times = []
                for t in times:
                    try:
                        hour_part = int(t.split(':')[0])
                        if hour_part < 12:
                            am_times.append(t)
                    except:
                        continue
                
                if am_times:
                    ws[f'B{current_row}'] = am_times[0]  # First AM time
                else:
                    ws[f'B{current_row}'] = ""
                
                # AM Departure (last time before 12:00)
                if am_times:
                    ws[f'C{current_row}'] = am_times[-1]  # Last AM time
                else:
                    ws[f'C{current_row}'] = ""
                
                # PM Arrival (first time 12:00 or after)
                pm_times = []
                for t in times:
                    try:
                        hour_part = int(t.split(':')[0])
                        if hour_part >= 12:
                            pm_times.append(t)
                    except:
                        continue
                
                if pm_times:
                    ws[f'D{current_row}'] = pm_times[0]  # First PM time
                else:
                    ws[f'D{current_row}'] = ""
                
                # PM Departure (last time 12:00 or after)
                if pm_times:
                    ws[f'E{current_row}'] = pm_times[-1]  # Last PM time
                else:
                    ws[f'E{current_row}'] = ""
                
                # Calculate undertime
                undertime_hours, undertime_minutes = calculate_undertime(
                    am_in=ws[f'B{current_row}'].value,
                    am_out=ws[f'C{current_row}'].value,
                    pm_in=ws[f'D{current_row}'].value,
                    pm_out=ws[f'E{current_row}'].value,
                    office_hours=office_hours
                )
                
                ws[f'F{current_row}'] = undertime_hours if undertime_hours else ""
                ws[f'G{current_row}'] = undertime_minutes if undertime_minutes else ""
                
                if undertime_hours:
                    total_undertime_hours += undertime_hours
                if undertime_minutes:
                    total_undertime_minutes += undertime_minutes
                
            else:
                # No data for this day
                for col in ['B', 'C', 'D', 'E', 'F', 'G']:
                    ws[f'{col}{current_row}'] = ""
            
            # Format cells
            for col in ['B', 'C', 'D', 'E']:
                cell = ws[f'{col}{current_row}']
                cell.font = Font(name='Arial', size=10, bold=True)
                cell.alignment = center_align
                cell.border = thin_border
                # Lock these cells (time cells should not be editable)
                cell.protection = Protection(locked=True)
            
            for col in ['F', 'G']:
                cell = ws[f'{col}{current_row}']
                cell.font = Font(name='Arial', size=10)
                cell.alignment = center_align
                cell.border = thin_border
        
        current_row += 1
    
    # ========== TOTAL UNDERTIME ROW ==========
    ws.merge_cells(f'A{current_row}:E{current_row}')
    ws[f'A{current_row}'] = "TOTAL"
    ws[f'A{current_row}'].font = Font(name='Arial', size=10, bold=True)
    ws[f'A{current_row}'].alignment = center_align
    ws[f'A{current_row}'].border = thin_border
    
    ws[f'F{current_row}'] = total_undertime_hours if total_undertime_hours else ""
    ws[f'F{current_row}'].font = Font(name='Arial', size=10, bold=True)
    ws[f'F{current_row}'].alignment = center_align
    ws[f'F{current_row}'].border = thin_border
    
    ws[f'G{current_row}'] = total_undertime_minutes if total_undertime_minutes else ""
    ws[f'G{current_row}'].font = Font(name='Arial', size=10, bold=True)
    ws[f'G{current_row}'].alignment = center_align
    ws[f'G{current_row}'].border = thin_border
    
    current_row += 2
    
    # ========== CERTIFICATION SECTION ==========
    ws.merge_cells(f'A{current_row}:G{current_row}')
    ws[f'A{current_row}'] = "I certify on my honor that the above is a true and correct report of the"
    ws[f'A{current_row}'].font = small_font
    ws[f'A{current_row}'].alignment = center_align
    
    current_row += 1
    
    ws.merge_cells(f'A{current_row}:G{current_row}')
    ws[f'A{current-row}'] = "hours of work performed, record of which was made daily at the time of"
    ws[f'A{current_row}'].font = small_font
    ws[f'A{current_row}'].alignment = center_align
    
    current_row += 1
    
    ws.merge_cells(f'A{current_row}:G{current_row}')
    ws[f'A{current_row}'] = "arrival and departure from office."
    ws[f'A{current_row}'].font = small_font
    ws[f'A{current_row}'].alignment = center_align
    
    current_row += 2
    
    # Signature lines
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = "_________________________________"
    ws[f'A{current_row}'].font = small_font
    ws[f'A{current_row}'].alignment = center_align
    
    ws.merge_cells(f'E{current_row}:G{current_row}')
    ws[f'E{current_row}'] = "_________________________________"
    ws[f'E{current_row}'].font = small_font
    ws[f'E{current_row}'].alignment = center_align
    
    current_row += 1
    
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = "Signature of Employee"
    ws[f'A{current_row}'].font = small_font
    ws[f'A{current_row}'].alignment = center_align
    
    ws.merge_cells(f'E{current_row}:G{current_row}')
    ws[f'E{current_row}'] = "Principal III"
    ws[f'E{current_row}'].font = small_font
    ws[f'E{current_row}'].alignment = center_align
    
    current_row += 1
    
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = ""
    
    ws.merge_cells(f'E{current_row}:G{current_row}')
    ws[f'E{current_row}'] = "VERIFIED as to the prescribed office hours:"
    ws[f'E{current_row}'].font = small_font
    ws[f'E{current_row}'].alignment = center_align
    
    # ========== ADJUST COLUMN WIDTHS ==========
    column_widths = {
        'A': 5,    # Day
        'B': 10,   # AM Arrival
        'C': 10,   # AM Departure
        'D': 10,   # PM Arrival
        'E': 10,   # PM Departure
        'F': 8,    # Undertime Hours
        'G': 8,    # Undertime Minutes
    }
    
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    # Protect the worksheet (except employee name and signature)
    ws.protection.sheet = True
    ws.protection.password = None  # No password, but prevents editing of time cells
    
    # Save to buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer

def calculate_undertime(am_in, am_out, pm_in, pm_out, office_hours):
    """Calculate undertime based on office hours"""
    
    # If any time is missing, return 0
    if not am_in or not am_out or not pm_in or not pm_out:
        return 0, 0
    
    try:
        # Parse times
        am_in_time = datetime.strptime(am_in, '%H:%M')
        am_out_time = datetime.strptime(am_out, '%H:%M')
        pm_in_time = datetime.strptime(pm_in, '%H:%M')
        pm_out_time = datetime.strptime(pm_out, '%H:%M')
        
        # Parse office hours
        office_am_in = datetime.strptime(office_hours['regular_am_in'], '%H:%M')
        office_am_out = datetime.strptime(office_hours['regular_am_out'], '%H:%M')
        office_pm_in = datetime.strptime(office_hours['regular_pm_in'], '%H:%M')
        office_pm_out = datetime.strptime(office_hours['regular_pm_out'], '%H:%M')
        
        # Calculate expected hours
        expected_am_hours = (office_am_out - office_am_in).seconds / 3600
        expected_pm_hours = (office_pm_out - office_pm_in).seconds / 3600
        total_expected_hours = expected_am_hours + expected_pm_hours
        
        # Calculate actual hours
        actual_am_hours = (am_out_time - am_in_time).seconds / 3600 if am_out_time > am_in_time else 0
        actual_pm_hours = (pm_out_time - pm_in_time).seconds / 3600 if pm_out_time > pm_in_time else 0
        total_actual_hours = actual_am_hours + actual_pm_hours
        
        # Calculate undertime
        undertime_hours_decimal = max(0, total_expected_hours - total_actual_hours)
        undertime_hours = int(undertime_hours_decimal)
        undertime_minutes = int((undertime_hours_decimal - undertime_hours) * 60)
        
        return undertime_hours, undertime_minutes
    
    except Exception as e:
        return 0, 0

def create_zip_file(excel_files, month, year):
    """Create ZIP file containing all DTR files"""
    
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_info in excel_files:
            # Clean filename - remove special characters
            clean_name = file_info['employee_name'].replace(',', '').replace('.', '').replace(' ', '_')
            filename = f"DTR_{clean_name}_{month}_{year}.xlsx"
            zip_file.writestr(filename, file_info['excel_file'].getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

# =============================================
# STREAMLIT APP - SIMPLE & ROBUST
# =============================================

# Page configuration
st.set_page_config(
    page_title="DTR Generator",
    page_icon="📋",
    layout="wide"
)

# Initialize session state
if 'raw_data' not in st.session_state:
    st.session_state.raw_data = None
if 'employee_settings' not in st.session_state:
    st.session_state.employee_settings = {}
if 'office_hours' not in st.session_state:
    st.session_state.office_hours = {
        'regular_am_in': '07:30',
        'regular_am_out': '11:50',
        'regular_pm_in': '12:50',
        'regular_pm_out': '16:30',
        'saturday': 'AS REQUIRED'
    }

# App title
st.title("📋 Civil Service Form No. 48 - DTR Generator")
st.markdown("Generate Daily Time Records from biometric attendance files")

# ========== FILE UPLOAD SECTION ==========
st.header("1. Upload Attendance File")

uploaded_file = st.file_uploader(
    "Choose your .dat file",
    type=['dat', 'txt', 'csv'],
    help="Upload biometric attendance file (ZKTeco format)"
)

if uploaded_file:
    try:
        # Read file as text
        content = uploaded_file.read().decode('utf-8', errors='ignore')
        lines = content.strip().split('\n')
        
        # Debug: Show file info
        st.info(f"📁 File: {uploaded_file.name} | Lines: {len(lines)}")
        
        # Parse ZKTeco .dat format
        data = []
        for i, line in enumerate(lines[:100]):  # Process first 100 lines for testing
            if line.strip():
                # ZKTeco format is usually: ID\tDateTime\tStatus
                parts = line.strip().split('\t')
                
                if len(parts) >= 2:
                    emp_no = parts[0].strip()
                    datetime_str = parts[1].strip()
                    
                    # Try to parse datetime
                    try:
                        # Common ZKTeco format: 2024-01-01 07:30:00
                        dt = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')
                        data.append({
                            'EmployeeNo': emp_no,
                            'DateTime': dt,
                            'Date': dt.date(),
                            'Time': dt.time(),
                            'Month': dt.month,
                            'Year': dt.year,
                            'Day': dt.day
                        })
                    except:
                        # Try other formats
                        try:
                            dt = datetime.strptime(datetime_str, '%m/%d/%Y %H:%M:%S')
                            data.append({
                                'EmployeeNo': emp_no,
                                'DateTime': dt,
                                'Date': dt.date(),
                                'Time': dt.time(),
                                'Month': dt.month,
                                'Year': dt.year,
                                'Day': dt.day
                            })
                        except:
                            # Skip if can't parse
                            continue
        
        if data:
            df = pd.DataFrame(data)
            st.session_state.raw_data = df
            
            st.success(f"✅ Successfully parsed {len(df)} records")
            
            # Show summary
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", len(df))
            with col2:
                st.metric("Unique Employees", df['EmployeeNo'].nunique())
            with col3:
                st.metric("Date Range", f"{df['Date'].min()} to {df['Date'].max()}")
            
            # Show sample
            with st.expander("👀 Preview Data"):
                st.dataframe(df.head(20))
        else:
            st.error("❌ Could not parse any valid records from the file.")
            
            # Show file content for debugging
            with st.expander("🔍 Debug - Show file content"):
                st.text("First 10 lines of file:")
                for i, line in enumerate(lines[:10]):
                    st.code(f"Line {i+1}: {line}")
    
    except Exception as e:
        st.error(f"❌ Error reading file: {str(e)}")

# ========== OFFICE HOURS SECTION ==========
st.header("2. Set Office Hours")

col1, col2 = st.columns(2)
with col1:
    am_in = st.text_input("AM Time In", "07:30", key="am_in")
    am_out = st.text_input("AM Time Out", "11:50", key="am_out")
with col2:
    pm_in = st.text_input("PM Time In", "12:50", key="pm_in")
    pm_out = st.text_input("PM Time Out", "16:30", key="pm_out")

saturday_hours = st.text_input("Saturday Hours", "AS REQUIRED", key="saturday")

st.session_state.office_hours = {
    'regular_am_in': am_in,
    'regular_am_out': am_out,
    'regular_pm_in': pm_in,
    'regular_pm_out': pm_out,
    'saturday': saturday_hours
}

# ========== MAIN PROCESSING SECTION ==========
if st.session_state.raw_data is not None:
    df = st.session_state.raw_data
    
    st.header("3. Generate DTR")
    
    # Month selection
    if 'Month' in df.columns and 'Year' in df.columns:
        unique_months = df[['Month', 'Year']].drop_duplicates().sort_values(['Year', 'Month'])
        
        if not unique_months.empty:
            # Create month options
            month_options = []
            for _, row in unique_months.iterrows():
                month_name = calendar.month_name[row['Month']]
                month_options.append(f"{month_name} {row['Year']}")
            
            selected_month = st.selectbox("Select Month", month_options)
            
            # Parse selection
            month_name, year_str = selected_month.split()
            month_num = list(calendar.month_name).index(month_name)
            year_num = int(year_str)
            
            # Filter data
            month_df = df[(df['Month'] == month_num) & (df['Year'] == year_num)].copy()
            
            if not month_df.empty:
                # Summary
                st.info(f"📊 **{len(month_df)} records found for {month_name} {year_num}**")
                
                # Get employees
                employees = sorted(month_df['EmployeeNo'].unique())
                
                # Initialize employee names
                for emp in employees:
                    if emp not in st.session_state.employee_settings:
                        st.session_state.employee_settings[emp] = {
                            'name': f"EMPLOYEE {emp}",
                            'employee_no': emp
                        }
                
                # Edit employee names
                with st.expander(f"✏️ Edit Employee Names ({len(employees)} employees)"):
                    for emp in employees:
                        current_name = st.session_state.employee_settings[emp]['name']
                        new_name = st.text_input(
                            f"Employee {emp}",
                            value=current_name,
                            key=f"name_{emp}"
                        )
                        if new_name != current_name:
                            st.session_state.employee_settings[emp]['name'] = new_name
                
                # Generate button
                st.markdown("---")
                
                if st.button("🚀 GENERATE DTR FILES NOW", type="primary", use_container_width=True):
                    with st.spinner(f"Generating {len(employees)} DTR files..."):
                        excel_files = []
                        errors = []
                        
                        for emp_no in employees:
                            try:
                                # Get employee data
                                emp_df = month_df[month_df['EmployeeNo'] == emp_no].copy()
                                
                                if emp_df.empty:
                                    errors.append(f"Employee {emp_no}: No records found")
                                    continue
                                
                                # Get employee name
                                emp_name = st.session_state.employee_settings.get(
                                    emp_no, 
                                    {'name': f"EMPLOYEE {emp_no}"}
                                )['name']
                                
                                # Generate DTR
                                excel_file = generate_civil_service_dtr(
                                    employee_no=emp_no,
                                    employee_name=emp_name,
                                    month=month_num,
                                    year=year_num,
                                    attendance_data=emp_df,
                                    office_hours=st.session_state.office_hours
                                )
                                
                                excel_files.append({
                                    'employee_no': emp_no,
                                    'employee_name': emp_name,
                                    'excel_file': excel_file
                                })
                                
                            except Exception as e:
                                errors.append(f"Employee {emp_no}: {str(e)[:100]}")
                                continue
                        
                        # Results
                        if excel_files:
                            st.success(f"✅ Successfully generated {len(excel_files)} DTR files!")
                            
                            if errors:
                                st.warning(f"⚠️ {len(errors)} errors occurred")
                                with st.expander("Show Errors"):
                                    for error in errors:
                                        st.write(error)
                            
                            # Create ZIP
                            zip_buffer = create_zip_file(excel_files, month_name, year_num)
                            
                            # Download buttons
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.download_button(
                                    label="📦 DOWNLOAD ALL (ZIP)",
                                    data=zip_buffer,
                                    file_name=f"DTR_{month_name}_{year_num}.zip",
                                    mime="application/zip",
                                    use_container_width=True
                                )
                            
                            with col2:
                                with st.expander("📥 Individual Files"):
                                    for file_info in excel_files:
                                        st.download_button(
                                            label=f"⬇️ {file_info['employee_name'][:20]}",
                                            data=file_info['excel_file'],
                                            file_name=f"DTR_{file_info['employee_name']}_{month_name}_{year_num}.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                        else:
                            st.error("❌ No files were generated.")
                            
                            if errors:
                                with st.expander("Error Details"):
                                    for error in errors:
                                        st.write(error)
            else:
                st.warning(f"No data found for {month_name} {year_num}")
    else:
        st.error("Dataframe doesn't have Month/Year columns. Please check file format.")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: gray;">
    <p>Civil Service Form No. 48 DTR Generator | Version 5.0 - Fixed File Reading</p>
    <p><small>Supports ZKTeco .dat format • Manual National High School</small></p>
    </div>
    """,
    unsafe_allow_html=True
)
