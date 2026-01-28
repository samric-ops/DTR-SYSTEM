import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, Protection
from io import BytesIO
import zipfile

# =============================================
# FIXED CIVIL SERVICE FORM NO. 48 GENERATOR
# =============================================

def generate_civil_service_dtr(employee_no, employee_name, month, year, attendance_data, office_hours):
    """Generate Civil Service Form No. 48 in Excel format"""
    
    # Create workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f"DTR {employee_name[:20]}"
    
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
    ws.merge_cells('A1:G1')
    ws['A1'] = "REPUBLIC OF THE PHILIPPINES"
    ws['A1'].font = header_font
    ws['A1'].alignment = center_align
    
    ws.merge_cells('A2:G2')
    ws['A2'] = "Department of Education"
    ws['A2'].font = header_font
    ws['A2'].alignment = center_align
    
    ws.merge_cells('A3:G3')
    ws['A3'] = "Division of Davao del Sur"
    ws['A3'].font = header_font
    ws['A3'].alignment = center_align
    
    ws.merge_cells('A4:G4')
    ws['A4'] = "Manual National High School"
    ws['A4'].font = header_font
    ws['A4'].alignment = center_align
    
    ws.merge_cells('A6:G6')
    ws['A6'] = "Civil Service Form No. 48"
    ws['A6'].font = small_font
    ws['A6'].alignment = center_align
    
    # FIX: Employee No. with proper formatting
    ws.merge_cells('A7:G7')
    ws['A7'] = f"Employee No. {employee_no}"
    ws['A7'].font = small_font
    ws['A7'].alignment = center_align
    
    # DAILY TIME RECORD Title
    ws.merge_cells('A9:G9')
    ws['A9'] = "DAILY TIME RECORD"
    ws['A9'].font = title_font
    ws['A9'].alignment = center_align
    
    ws.merge_cells('A10:G10')
    ws['A10'] = "-------------------------o0o-------------------------"
    ws['A10'].alignment = center_align
    
    # Employee Name - FIX: Proper name display
    ws.merge_cells('A12:G12')
    ws['A12'] = employee_name.upper()
    ws['A12'].font = Font(name='Arial', size=12, bold=True)
    ws['A12'].alignment = center_align
    
    ws.merge_cells('A13:G13')
    ws['A13'] = "(Name)"
    ws['A13'].font = small_font
    ws['A13'].alignment = center_align
    
    # ========== MONTH AND OFFICE HOURS SECTION ==========
    current_row = 15
    
    month_name = calendar.month_name[month]
    ws.merge_cells(f'A{current_row}:C{current_row}')
    ws[f'A{current_row}'] = f"{month_name.upper()} {year}"
    ws[f'A{current_row}'].font = Font(name='Arial', size=11, bold=True)
    ws[f'A{current_row}'].alignment = center_align
    
    ws.merge_cells(f'E{current_row}:G{current_row}')
    ws[f'E{current_row}'] = "For the month of"
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
    
    current_row += 2
    
    # ========== DTR TABLE HEADER ==========
    ws.merge_cells(f'A{current_row}:A{current_row+1}')
    ws[f'A{current_row}'] = "Day"
    ws[f'A{current_row}'].font = header_font
    ws[f'A{current_row}'].alignment = center_align
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:C{current_row}')
    ws[f'B{current_row}'] = "A.M."
    ws[f'B{current_row}'].font = header_font
    ws[f'B{current_row}'].alignment = center_align
    ws[f'B{current_row}'].border = thin_border
    
    ws.merge_cells(f'D{current_row}:E{current_row}')
    ws[f'D{current_row}'] = "P.M."
    ws[f'D{current_row}'].font = header_font
    ws[f'D{current_row}'].alignment = center_align
    ws[f'D{current_row}'].border = thin_border
    
    ws.merge_cells(f'F{current_row}:G{current_row}')
    ws[f'F{current_row}'] = "Undertime"
    ws[f'F{current_row}'].font = header_font
    ws[f'F{current_row}'].alignment = center_align
    ws[f'F{current_row}'].border = thin_border
    
    current_row += 1
    
    subheader_cols = ['B', 'C', 'D', 'E', 'F', 'G']
    subheader_texts = ['Arrival', 'Departure', 'Arrival', 'Departure', 'Hours', 'Minutes']
    
    for col, text in zip(subheader_cols, subheader_texts):
        ws[f'{col}{current_row}'] = text
        ws[f'{col}{current_row}'].font = small_font
        ws[f'{col}{current_row}'].alignment = center_align
        ws[f'{col}{current_row}'].border = thin_border
    
    current_row += 1
    
    # ========== DTR TABLE DATA ==========
    days_in_month = calendar.monthrange(year, month)[1]
    
    # Process attendance data
    attendance_by_day = {}
    if attendance_data is not None and not attendance_data.empty:
        for _, row in attendance_data.iterrows():
            try:
                day = int(row['Day'])
                time_val = row['Time']
                
                if hasattr(time_val, 'strftime'):
                    time_str = time_val.strftime('%H:%M')
                elif isinstance(time_val, str):
                    time_str = time_val
                elif hasattr(time_val, '__str__'):
                    time_str = str(time_val)
                else:
                    continue
                
                if day not in attendance_by_day:
                    attendance_by_day[day] = []
                attendance_by_day[day].append(time_str)
            except:
                continue
    
    total_undertime_hours = 0
    total_undertime_minutes = 0
    
    for day in range(1, days_in_month + 1):
        date_obj = datetime(year, month, day)
        day_name = date_obj.strftime('%A')
        
        # Day cell
        ws[f'A{current_row}'] = str(day)
        ws[f'A{current_row}'].font = Font(name='Arial', size=10, bold=True)
        ws[f'A{current_row}'].alignment = center_align
        ws[f'A{current_row}'].border = thin_border
        
        if day_name.upper() == 'SATURDAY':
            ws.merge_cells(f'B{current_row}:C{current_row}')
            ws[f'B{current_row}'] = "SATURDAY"
            ws[f'B{current_row}'].font = Font(name='Arial', size=10, bold=True, italic=True)
            ws[f'B{current_row}'].alignment = center_align
            ws[f'B{current_row}'].border = thin_border
            ws[f'B{current_row}'].fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
            
            for col in ['D', 'E', 'F', 'G']:
                ws[f'{col}{current_row}'].border = thin_border
        
        elif day_name.upper() == 'SUNDAY':
            ws.merge_cells(f'B{current_row}:C{current_row}')
            ws[f'B{current_row}'] = "SUNDAY"
            ws[f'B{current_row}'].font = Font(name='Arial', size=10, bold=True, italic=True)
            ws[f'B{current_row}'].alignment = center_align
            ws[f'B{current_row}'].border = thin_border
            ws[f'B{current_row}'].fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
            
            for col in ['D', 'E', 'F', 'G']:
                ws[f'{col}{current_row}'].border = thin_border
        
        else:
            # Regular work day
            if day in attendance_by_day:
                times = sorted(attendance_by_day[day])
                
                # AM times (before 12:00)
                am_times = []
                for t in times:
                    try:
                        hour_part = int(t.split(':')[0])
                        if hour_part < 12:
                            am_times.append(t)
                    except:
                        continue
                
                if am_times:
                    ws[f'B{current_row}'] = am_times[0]
                    ws[f'C{current_row}'] = am_times[-1]
                else:
                    ws[f'B{current_row}'] = ""
                    ws[f'C{current_row}'] = ""
                
                # PM times (12:00 and after)
                pm_times = []
                for t in times:
                    try:
                        hour_part = int(t.split(':')[0])
                        if hour_part >= 12:
                            pm_times.append(t)
                    except:
                        continue
                
                if pm_times:
                    ws[f'D{current_row}'] = pm_times[0]
                    ws[f'E{current_row}'] = pm_times[-1]
                else:
                    ws[f'D{current_row}'] = ""
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
                for col in ['B', 'C', 'D', 'E', 'F', 'G']:
                    ws[f'{col}{current_row}'] = ""
            
            # Format cells
            for col in ['B', 'C', 'D', 'E']:
                cell = ws[f'{col}{current_row}']
                cell.font = Font(name='Arial', size=10, bold=True)
                cell.alignment = center_align
                cell.border = thin_border
            
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
    
    # Save to buffer
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer

def calculate_undertime(am_in, am_out, pm_in, pm_out, office_hours):
    """Calculate undertime based on office hours"""
    if not am_in or not am_out or not pm_in or not pm_out:
        return 0, 0
    
    try:
        am_in_time = datetime.strptime(am_in, '%H:%M')
        am_out_time = datetime.strptime(am_out, '%H:%M')
        pm_in_time = datetime.strptime(pm_in, '%H:%M')
        pm_out_time = datetime.strptime(pm_out, '%H:%M')
        
        office_am_in = datetime.strptime(office_hours['regular_am_in'], '%H:%M')
        office_am_out = datetime.strptime(office_hours['regular_am_out'], '%H:%M')
        office_pm_in = datetime.strptime(office_hours['regular_pm_in'], '%H:%M')
        office_pm_out = datetime.strptime(office_hours['regular_pm_out'], '%H:%M')
        
        expected_am = (office_am_out - office_am_in).seconds / 3600
        expected_pm = (office_pm_out - office_pm_in).seconds / 3600
        total_expected = expected_am + expected_pm
        
        actual_am = (am_out_time - am_in_time).seconds / 3600 if am_out_time > am_in_time else 0
        actual_pm = (pm_out_time - pm_in_time).seconds / 3600 if pm_out_time > pm_in_time else 0
        total_actual = actual_am + actual_pm
        
        undertime_decimal = max(0, total_expected - total_actual)
        undertime_hours = int(undertime_decimal)
        undertime_minutes = int((undertime_decimal - undertime_hours) * 60)
        
        return undertime_hours, undertime_minutes
    except:
        return 0, 0

def create_zip_file(excel_files, month, year):
    """Create ZIP file containing all DTR files"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_info in excel_files:
            clean_name = file_info['employee_name'].replace(',', '').replace('.', '').replace(' ', '_')
            filename = f"DTR_{clean_name}_{month}_{year}.xlsx"
            zip_file.writestr(filename, file_info['excel_file'].getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

# =============================================
# STREAMLIT APP - IMPROVED VERSION
# =============================================

# Page configuration
st.set_page_config(
    page_title="DTR Generator - Manual NHS",
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

# App title with better styling
st.markdown("""
    <h1 style='text-align: center; color: #1E3A8A; margin-bottom: 20px;'>
        📋 Civil Service Form No. 48 - DTR Generator
    </h1>
""", unsafe_allow_html=True)

st.markdown("""
    <p style='text-align: center; color: #4B5563; margin-bottom: 30px;'>
        Manual National High School - Division of Davao del Sur
    </p>
""", unsafe_allow_html=True)

# ========== FILE UPLOAD SECTION ==========
with st.expander("📤 **1. UPLOAD ATTENDANCE FILE**", expanded=True):
    uploaded_file = st.file_uploader(
        "Choose biometric attendance file (.dat, .txt, .csv)",
        type=['dat', 'txt', 'csv'],
        help="Upload the file exported from your biometric system"
    )
    
    if uploaded_file:
        try:
            # Read and parse file
            content = uploaded_file.read().decode('utf-8', errors='ignore')
            lines = [line.strip() for line in content.split('\n') if line.strip()]
            
            data = []
            for line in lines:
                parts = line.split()
                if len(parts) >= 2:
                    emp_no = parts[0].strip()
                    datetime_str = ' '.join(parts[1:3])
                    
                    # Try different datetime formats
                    date_formats = [
                        '%Y-%m-%d %H:%M:%S',
                        '%m/%d/%Y %H:%M:%S',
                        '%d/%m/%Y %H:%M:%S',
                        '%Y/%m/%d %H:%M:%S'
                    ]
                    
                    dt = None
                    for fmt in date_formats:
                        try:
                            dt = datetime.strptime(datetime_str, fmt)
                            break
                        except:
                            continue
                    
                    if dt:
                        data.append({
                            'EmployeeNo': emp_no,
                            'DateTime': dt,
                            'Date': dt.date(),
                            'Time': dt.time(),
                            'Month': dt.month,
                            'Year': dt.year,
                            'Day': dt.day
                        })
            
            if data:
                df = pd.DataFrame(data)
                st.session_state.raw_data = df
                
                st.success(f"✅ Successfully loaded {len(df)} records")
                
                # Show summary
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Records", len(df))
                with col2:
                    st.metric("Unique Employees", df['EmployeeNo'].nunique())
                with col3:
                    st.metric("Date Range", f"{df['Date'].min()} to {df['Date'].max()}")
                
                # Show preview
                with st.expander("👀 Preview Data"):
                    st.dataframe(df.head(10))
            else:
                st.error("❌ No valid records found in the file.")
                
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")

# ========== OFFICE HOURS SECTION ==========
with st.expander("⏰ **2. SET OFFICE HOURS**", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Morning Session")
        am_in = st.text_input("AM Time In", "07:30", key="am_in")
        am_out = st.text_input("AM Time Out", "11:50", key="am_out")
    with col2:
        st.subheader("Afternoon Session")
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
    
    with st.expander("🚀 **3. GENERATE DTR FILES**", expanded=True):
        # Month selection
        if 'Month' in df.columns and 'Year' in df.columns:
            unique_months = df[['Month', 'Year']].drop_duplicates().sort_values(['Year', 'Month'])
            
            if not unique_months.empty:
                # Create month options
                month_options = []
                for _, row in unique_months.iterrows():
                    month_name = calendar.month_name[row['Month']]
                    month_options.append(f"{month_name} {row['Year']}")
                
                selected_month = st.selectbox("Select Month to Generate DTR", month_options)
                
                # Parse selection
                month_name, year_str = selected_month.split()
                month_num = list(calendar.month_name).index(month_name)
                year_num = int(year_str)
                
                # Filter data for selected month
                month_df = df[(df['Month'] == month_num) & (df['Year'] == year_num)].copy()
                
                if not month_df.empty:
                    # Summary
                    st.info(f"📊 **Found {len(month_df)} records for {month_name} {year_num}**")
                    
                    # Get employees in this month
                    employees_in_month = sorted(month_df['EmployeeNo'].unique())
                    
                    # FIX: Pre-load specific employees with names
                    # Add Richard P. Samoranos with employee no. 7220970
                    if '7220970' not in st.session_state.employee_settings:
                        st.session_state.employee_settings['7220970'] = {
                            'name': 'RICHARD P. SAMORANOS',
                            'employee_no': '7220970'
                        }
                    
                    # Employee names editor
                    st.subheader("✏️ Edit Employee Names")
                    st.write("Enter the correct name for each biometric ID:")
                    
                    # Create a form for better management
                    name_form = st.form(key='employee_names_form')
                    
                    employee_names = {}
                    with name_form:
                        cols = st.columns(2)
                        for idx, emp_id in enumerate(employees_in_month):
                            col_idx = idx % 2
                            with cols[col_idx]:
                                # Get current name
                                current_name = st.session_state.employee_settings.get(
                                    emp_id, 
                                    {'name': f"EMPLOYEE {emp_id}"}
                                )['name']
                                
                                # Edit field with label
                                new_name = st.text_input(
                                    f"**ID: {emp_id}**",
                                    value=current_name,
                                    key=f"name_{emp_id}",
                                    help=f"Enter name for employee with biometric ID {emp_id}"
                                )
                                
                                employee_names[emp_id] = new_name
        
                        # Save all names at once
                        if st.form_submit_button("💾 SAVE ALL NAMES"):
                            for emp_id, emp_name in employee_names.items():
                                st.session_state.employee_settings[emp_id] = {
                                    'name': emp_name,
                                    'employee_no': emp_id
                                }
                            st.success("✅ All names saved successfully!")
                    
                    # Generate button
                    st.markdown("---")
                    
                    if st.button("🚀 GENERATE DTR FILES NOW", type="primary", use_container_width=True):
                        with st.spinner(f"Generating DTR files for {len(employees_in_month)} employees..."):
                            progress_bar = st.progress(0)
                            excel_files = []
                            success_count = 0
                            error_list = []
                            
                            for idx, emp_id in enumerate(employees_in_month):
                                try:
                                    # Get employee data
                                    emp_df = month_df[month_df['EmployeeNo'] == emp_id].copy()
                                    
                                    if emp_df.empty:
                                        error_list.append(f"❌ {emp_id}: No attendance data")
                                        continue
                                    
                                    # Get employee name
                                    emp_settings = st.session_state.employee_settings.get(
                                        emp_id, 
                                        {'name': f"EMPLOYEE {emp_id}", 'employee_no': emp_id}
                                    )
                                    emp_name = emp_settings['name']
                                    
                                    # Generate DTR
                                    excel_file = generate_civil_service_dtr(
                                        employee_no=emp_id,
                                        employee_name=emp_name,
                                        month=month_num,
                                        year=year_num,
                                        attendance_data=emp_df,
                                        office_hours=st.session_state.office_hours
                                    )
                                    
                                    excel_files.append({
                                        'employee_no': emp_id,
                                        'employee_name': emp_name,
                                        'excel_file': excel_file
                                    })
                                    
                                    success_count += 1
                                    
                                except Exception as e:
                                    error_list.append(f"❌ {emp_id}: {str(e)[:100]}")
                                
                                # Update progress
                                progress_bar.progress((idx + 1) / len(employees_in_month))
                            
                            # Show results
                            if excel_files:
                                st.success(f"✅ Successfully generated {success_count} DTR files!")
                                
                                # Create ZIP
                                zip_buffer = create_zip_file(excel_files, month_name, year_num)
                                
                                # Download buttons
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    st.download_button(
                                        label="📦 DOWNLOAD ALL FILES (ZIP)",
                                        data=zip_buffer,
                                        file_name=f"DTR_{month_name}_{year_num}.zip",
                                        mime="application/zip",
                                        use_container_width=True,
                                        help="Download all DTR files in a single ZIP archive"
                                    )
                                
                                with col2:
                                    with st.expander("📥 Download Individual Files"):
                                        for file_info in excel_files:
                                            # Special styling for Richard P. Samoranos
                                            if file_info['employee_no'] == '7220970':
                                                st.success(f"✅ **{file_info['employee_name']}** (ID: {file_info['employee_no']})")
                                            else:
                                                st.write(f"{file_info['employee_name']} (ID: {file_info['employee_no']})")
                                            
                                            st.download_button(
                                                label=f"⬇️ Download {file_info['employee_name'][:20]}",
                                                data=file_info['excel_file'],
                                                file_name=f"DTR_{file_info['employee_name']}_{month_name}_{year_num}.xlsx",
                                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                key=f"download_{file_info['employee_no']}"
                                            )
                                
                                # Show errors if any
                                if error_list:
                                    with st.expander("⚠️ View Errors"):
                                        for error in error_list:
                                            st.write(error)
                            else:
                                st.error("❌ No files were generated.")
                else:
                    st.warning(f"No data found for {month_name} {year_num}")
        else:
            st.error("Data doesn't have Month/Year columns.")

# Footer with better styling
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #6B7280; padding: 20px;">
        <p><strong>Civil Service Form No. 48 DTR Generator</strong> | Version 8.0</p>
        <p><small>Manual National High School - Division of Davao del Sur</small></p>
    </div>
""", unsafe_allow_html=True)

# Sidebar for additional info
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/9/99/DepEd_seal.svg/1200px-DepEd_seal.svg.png", 
             width=100)
    st.title("Quick Guide")
    
    st.markdown("""
    ### How to Use:
    1. **Upload** biometric attendance file (.dat/.txt/.csv)
    2. **Set** office hours for regular days
    3. **Edit** employee names for each biometric ID
    4. **Select** month to generate
    5. **Generate** and download DTR files
    
    ### Employee No. 7220970:
    - This ID belongs to **Richard P. Samoranos**
    - The name will be pre-loaded automatically
    
    ### File Format:
    Ensure your attendance file has:
    - Employee No. (first column)
    - Date and Time (YYYY-MM-DD HH:MM:SS)
    - One record per line
    """)
