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
    
    # Process attendance data by day
    attendance_by_day = {}
    for _, row in attendance_data.iterrows():
        day = row['Day']
        time_str = row['Time'].strftime('%H:%M')
        
        if day not in attendance_by_day:
            attendance_by_day[day] = []
        attendance_by_day[day].append(time_str)
    
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
                am_times = [t for t in times if int(t.split(':')[0]) < 12]
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
                pm_times = [t for t in times if int(t.split(':')[0]) >= 12]
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
    ws[f'A{current_row}'] = "hours of work performed, record of which was made daily at the time of"
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
            filename = f"DTR_{file_info['employee_name']}_{month}_{year}.xlsx"
            zip_file.writestr(filename, file_info['excel_file'].getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

# =============================================
# STREAMLIT APP STARTS HERE
# =============================================

# Page configuration
st.set_page_config(
    page_title="Civil Service Form No. 48 DTR Generator",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 1rem;
    }
    .cs-form {
        background-color: #F3F4F6;
        padding: 2rem;
        border-radius: 10px;
        border: 2px solid #1E3A8A;
        margin: 1rem 0;
    }
    .form-title {
        text-align: center;
        font-weight: bold;
        color: #1E3A8A;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# App title
st.markdown('<h1 class="main-header">📋 Civil Service Form No. 48 - DTR Generator</h1>', unsafe_allow_html=True)
st.markdown("Generate Daily Time Records from .dat files following DepEd format")

# Initialize session state
if 'dtr_data' not in st.session_state:
    st.session_state.dtr_data = {}
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

# Sidebar for settings
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3131/3131639.png", width=100)
    st.title("⚙️ Settings")
    
    st.markdown("---")
    
    # Step 1: Upload DAT file
    st.subheader("1. Upload Attendance Data")
    uploaded_file = st.file_uploader(
        "Choose .dat file",
        type=['dat', 'txt', 'csv'],
        help="Upload your biometric attendance file"
    )
    
    if uploaded_file:
        try:
            # Read the DAT file
            content = uploaded_file.getvalue().decode('utf-8')
            first_line = content.split('\n')[0].strip()
            
            if '\t' in first_line:
                delimiter = '\t'
            elif ',' in first_line:
                delimiter = ','
            else:
                delimiter = ' '
            
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, header=None, delimiter=delimiter)
            
            # Check number of columns and assign names
            if len(df.columns) >= 2:
                df.columns = ['EmployeeNo', 'DateTime'] + [f'Col{i}' for i in range(2, len(df.columns))]
                
                # Convert DateTime
                df['DateTime'] = pd.to_datetime(df['DateTime'], errors='coerce')
                df = df.dropna(subset=['DateTime'])
                
                # Add date components
                df['Date'] = df['DateTime'].dt.date
                df['Time'] = df['DateTime'].dt.time
                df['Month'] = df['DateTime'].dt.month
                df['Year'] = df['DateTime'].dt.year
                df['Day'] = df['DateTime'].dt.day
                df['DayName'] = df['DateTime'].dt.day_name()
                df['Hour'] = df['DateTime'].dt.hour
                df['Minute'] = df['DateTime'].dt.minute
                
                # Store in session state
                st.session_state.raw_data = df
                
                # Show summary
                st.success(f"✅ File loaded: {len(df)} records")
                st.info(f"**Employees:** {df['EmployeeNo'].nunique()}")
                st.info(f"**Date Range:** {df['Date'].min()} to {df['Date'].max()}")
                
                # Extract unique months and years
                months_data = df[['Month', 'Year']].drop_duplicates().sort_values(['Year', 'Month'])
                st.session_state.available_months = months_data.to_dict('records')
                
            else:
                st.error("Invalid file format. Need at least 2 columns.")
                
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
    
    st.markdown("---")
    
    # Step 2: Office Hours Settings
    st.subheader("2. Set Office Hours")
    
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.office_hours['regular_am_in'] = st.text_input(
            "AM Time In", 
            value=st.session_state.office_hours['regular_am_in'],
            help="e.g., 07:30"
        )
    with col2:
        st.session_state.office_hours['regular_am_out'] = st.text_input(
            "AM Time Out", 
            value=st.session_state.office_hours['regular_am_out'],
            help="e.g., 11:50"
        )
    
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.office_hours['regular_pm_in'] = st.text_input(
            "PM Time In", 
            value=st.session_state.office_hours['regular_pm_in'],
            help="e.g., 12:50"
        )
    with col2:
        st.session_state.office_hours['regular_pm_out'] = st.text_input(
            "PM Time Out", 
            value=st.session_state.office_hours['regular_pm_out'],
            help="e.g., 16:30"
        )
    
    st.session_state.office_hours['saturday'] = st.text_input(
        "Saturday Hours",
        value=st.session_state.office_hours['saturday'],
        help="Usually 'AS REQUIRED'"
    )
    
    st.markdown("---")
    
    # Step 3: Employee Settings
    if 'raw_data' in st.session_state:
        st.subheader("3. Employee Settings")
        
        employees = sorted(st.session_state.raw_data['EmployeeNo'].unique())
        selected_employee = st.selectbox("Select Employee", employees)
        
        # Employee name input
        employee_name = st.text_input(
            f"Full Name for Employee {selected_employee}",
            value=f"EMPLOYEE {selected_employee}",
            key=f"name_{selected_employee}"
        )
        
        # Store in session state
        if selected_employee not in st.session_state.employee_settings:
            st.session_state.employee_settings[selected_employee] = {
                'name': employee_name,
                'employee_no': selected_employee
            }
        else:
            st.session_state.employee_settings[selected_employee]['name'] = employee_name

# Main Content
if 'raw_data' in st.session_state:
    df = st.session_state.raw_data
    
    # Month Selection
    st.markdown("---")
    st.subheader("📅 Select Month for DTR")
    
    # Create month-year options
    month_options = []
    for month_data in st.session_state.available_months:
        month_num = month_data['Month']
        year_num = month_data['Year']
        month_name = calendar.month_name[month_num]
        month_options.append(f"{month_name} {year_num}")
    
    if month_options:
        selected_period = st.selectbox("Choose Month", month_options)
        
        # Parse selected month and year
        month_name, year_str = selected_period.split()
        month_num = list(calendar.month_name).index(month_name)
        year_num = int(year_str)
        
        # Filter data for selected month
        month_df = df[(df['Month'] == month_num) & (df['Year'] == year_num)]
        
        if not month_df.empty:
            # Display summary
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Month", f"{month_name} {year_num}")
            with col2:
                st.metric("Total Records", len(month_df))
            with col3:
                st.metric("Employees", month_df['EmployeeNo'].nunique())
            with col4:
                days = month_df['Date'].nunique()
                st.metric("Days with Data", days)
            
            # Generate DTR Button
            st.markdown("---")
            st.subheader("🔄 Generate DTR Files")
            
            if st.button("📋 Generate Civil Service Form No. 48", type="primary", use_container_width=True):
                with st.spinner("Generating DTR files..."):
                    try:
                        # Generate DTR for each employee
                        employees = sorted(month_df['EmployeeNo'].unique())
                        excel_files = []
                        
                        for emp_no in employees:
                            # Get employee data
                            emp_df = month_df[month_df['EmployeeNo'] == emp_no].copy()
                            
                            # Get employee name from settings
                            emp_name = st.session_state.employee_settings.get(
                                str(emp_no), 
                                {'name': f"EMPLOYEE {emp_no}"}
                            )['name']
                            
                            # Generate DTR Excel
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
                                'excel_file': excel_file,
                                'month': month_name,
                                'year': year_num
                            })
                        
                        # Create ZIP file
                        zip_buffer = create_zip_file(excel_files, month_name, year_num)
                        
                        # Success message
                        st.success(f"✅ Generated {len(excel_files)} DTR files for {month_name} {year_num}")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download All DTR Files (ZIP)",
                            data=zip_buffer,
                            file_name=f"DTR_CivilService_{month_name}_{year_num}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                        
                        # Individual download buttons
                        st.subheader("📥 Download Individual DTR")
                        cols = st.columns(3)
                        for idx, file_info in enumerate(excel_files):
                            with cols[idx % 3]:
                                st.download_button(
                                    label=f"{file_info['employee_name']}",
                                    data=file_info['excel_file'],
                                    file_name=f"DTR_{file_info['employee_name']}_{month_name}_{year_num}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        # Save data option
                        st.markdown("---")
                        st.subheader("💾 Save Data for Future Printing")
                        
                        save_name = st.text_input(
                            "Save as (optional):",
                            value=f"DTR_Data_{month_name}_{year_num}",
                            help="This will save the processed data for future re-printing"
                        )
                        
                        if st.button("💾 Save Data for This Month"):
                            # Save to session state
                            key = f"{month_name}_{year_num}"
                            st.session_state.dtr_data[key] = {
                                'month_df': month_df,
                                'office_hours': st.session_state.office_hours,
                                'employee_settings': st.session_state.employee_settings,
                                'timestamp': datetime.now()
                            }
                            st.success(f"✅ Data saved for {month_name} {year_num}")
                    
                    except Exception as e:
                        st.error(f"Error generating DTR: {str(e)}")
        
        else:
            st.warning(f"No data found for {month_name} {year_num}")
    
    else:
        st.warning("No complete months found in the data")
    
    # Previous Saved Data Section
    st.markdown("---")
    st.subheader("📚 Previously Saved DTR Data")
    
    if st.session_state.dtr_data:
        for key, data in st.session_state.dtr_data.items():
            with st.expander(f"📅 {key} - Saved {data['timestamp'].strftime('%Y-%m-%d %H:%M')}"):
                st.write(f"**Employees:** {len(data['month_df']['EmployeeNo'].unique())}")
                st.write(f"**Records:** {len(data['month_df'])}")
                st.write(f"**Office Hours:** {data['office_hours']}")
                
                if st.button(f"🔄 Re-generate DTR for {key}", key=f"regenerate_{key}"):
                    # Re-generate DTR from saved data
                    pass
    
    # Data Preview
    st.markdown("---")
    st.subheader("👀 Data Preview")
    
    tab1, tab2 = st.tabs(["Raw Data", "Daily Summary"])
    
    with tab1:
        st.dataframe(df.head(100), use_container_width=True)
    
    with tab2:
        if 'month_df' in locals():
            # Create daily summary
            summary = []
            for date in sorted(month_df['Date'].unique()):
                daily = month_df[month_df['Date'] == date]
                for emp in daily['EmployeeNo'].unique():
                    emp_daily = daily[daily['EmployeeNo'] == emp]
                    time_in = emp_daily['Time'].min()
                    time_out = emp_daily['Time'].max()
                    
                    summary.append({
                        'Date': date,
                        'Employee': emp,
                        'Records': len(emp_daily),
                        'First Log': time_in,
                        'Last Log': time_out
                    })
            
            if summary:
                summary_df = pd.DataFrame(summary)
                st.dataframe(summary_df, use_container_width=True)

else:
    # Welcome screen
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="cs-form">
        <div class="form-title">Civil Service Form No. 48</div>
        <h3>📋 DAILY TIME RECORD</h3>
        <p><strong>Department of Education</strong></p>
        <p>Division of Davao del Sur</p>
        <p>Manual National High School</p>
        <hr>
        <p><strong>How to use:</strong></p>
        <ol>
            <li>Upload your .dat file (biometric data)</li>
            <li>Set office hours in settings</li>
            <li>Enter employee names</li>
            <li>Select month to generate DTR</li>
            <li>Download Excel files</li>
        </ol>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div style="background-color: #EFF6FF; padding: 2rem; border-radius: 10px;">
        <h3>🎯 Features</h3>
        <ul>
            <li>✅ <strong>Exact Civil Service Form No. 48 format</strong></li>
            <li>✅ <strong>Auto-fill from .dat files</strong></li>
            <li>✅ <strong>Employee No. from .dat file</strong></li>
            <li>✅ <strong>Month/Year from data</strong></li>
            <li>✅ <strong>Save data for future printing</strong></li>
            <li>✅ <strong>Customizable office hours</strong></li>
            <li>✅ <strong>Auto-mark Saturdays/Sundays</strong></li>
            <li>✅ <strong>Protected time cells</strong></li>
            <li>✅ <strong>Excel format</strong></li>
        </ul>
        
        <h3>📝 Required Format</h3>
        <p>.dat file should have:</p>
        <pre>
        EmployeeNo,DateTime
        7220970,2026-01-05 06:51:00
        7220970,2026-01-05 11:54:00
        </pre>
        </div>
        """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: gray;">
    <p>Civil Service Form No. 48 DTR Generator v2.0 | For DepEd Manual National High School</p>
    </div>
    """,
    unsafe_allow_html=True
)
