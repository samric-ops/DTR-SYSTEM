import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar
import zipfile
from io import BytesIO
import re

# =============================================
# PDF DTR GENERATOR - EXACT TEMPLATE MATCH
# =============================================

try:
    from fpdf import FPDF
    
    class DTR_PDF(FPDF):
        def __init__(self):
            super().__init__(format='A4')  # A4 paper size
            self.set_auto_page_break(auto=False)  # Manual page breaks
            self.left_margin = 15
            self.right_margin = 15
            self.set_margins(self.left_margin, 25, self.right_margin)
            
        def header(self):
            # This is empty because we'll manually create everything
            pass
            
        def create_dtr_template(self, employee_no, employee_name, month_year, office_hours, attendance_by_day, month_num, year_num, office_hours_dict):
            """Create DTR following EXACT template"""
            
            # ========== HEADER SECTION ==========
            # Set Y position for header
            self.set_y(20)
            
            # REPUBLIC OF THE PHILIPPINES
            self.set_font("Arial", "B", 11)
            self.cell(0, 5, "REPUBLIC OF THE PHILIPPINES", 0, 1, "L")
            
            # Department of Education
            self.set_font("Arial", "B", 10)
            self.cell(0, 5, "Department of Education", 0, 1, "L")
            
            # Division of Davao del Sur
            self.cell(0, 5, "Division of Davao del Sur", 0, 1, "L")
            
            # Manual National High School
            self.cell(0, 5, "Manual National High School", 0, 1, "L")
            
            # Move right for Civil Service Form No. 48
            self.set_y(20)
            self.set_x(140)
            self.set_font("Arial", "I", 9)
            self.cell(0, 5, "Civil Service Form No. 48", 0, 1, "C")
            
            # Employee No.
            self.set_y(25)
            self.set_x(140)
            self.set_font("Arial", "", 9)
            self.cell(0, 5, f"Employee No. {employee_no}", 0, 1, "C")
            
            # ========== DAILY TIME RECORD TITLE ==========
            self.set_y(45)
            self.set_font("Arial", "B", 14)
            self.cell(0, 8, "DAILY TIME RECORD", 0, 1, "C")
            
            # Separator line
            self.set_font("Arial", "", 10)
            self.cell(0, 5, "-------------------------o0o-------------------------", 0, 1, "C")
            
            # Employee Name
            self.set_y(65)
            self.set_font("Arial", "B", 12)
            self.cell(0, 6, employee_name.upper(), 0, 1, "C")
            
            # (Name) label
            self.set_font("Arial", "", 9)
            self.cell(0, 4, "(Name)", 0, 1, "C")
            
            # ========== MONTH AND OFFICE HOURS ==========
            # Left side: Month and Year
            self.set_y(90)
            self.set_font("Arial", "B", 11)
            self.cell(50, 6, month_year.upper(), 0, 0, "L")
            
            # Left side: Office hours
            self.set_y(100)
            self.set_font("Arial", "", 10)
            
            # AM Hours
            self.cell(50, 6, office_hours['am'], 0, 0, "L")
            self.ln(6)
            
            # PM Hours
            self.set_x(self.left_margin)
            self.cell(50, 6, office_hours['pm'], 0, 0, "L")
            self.ln(6)
            
            # Saturday Hours
            self.set_x(self.left_margin)
            self.cell(50, 6, office_hours['saturday'], 0, 0, "L")
            
            # Right side: Labels
            self.set_y(90)
            self.set_x(140)
            self.set_font("Arial", "I", 9)
            self.cell(50, 6, "For the month of", 0, 2, "L")
            
            self.set_x(140)
            self.cell(50, 6, "Official hours for arrival", 0, 2, "L")
            self.set_x(140)
            self.cell(50, 6, "and departure", 0, 2, "L")
            
            self.set_x(140)
            self.ln(8)
            self.set_font("Arial", "I", 9)
            self.cell(30, 6, "Regular days", 0, 2, "L")
            self.set_x(140)
            self.cell(30, 6, "Saturdays", 0, 2, "L")
            
            # ========== DTR TABLE ==========
            self.set_y(140)
            self.create_dtr_table(attendance_by_day, month_num, year_num, office_hours_dict)
            
            # ========== CERTIFICATION SECTION ==========
            # Get Y position after table
            current_y = self.get_y()
            self.set_y(current_y + 10)
            
            # Certification text
            self.set_font("Arial", "", 9)
            self.cell(0, 4, "I certify on my honor that the above is a true and correct report of the", 0, 1, "C")
            self.cell(0, 4, "hours of work performed, record of which was made daily at the time of", 0, 1, "C")
            self.cell(0, 4, "arrival and departure from office.", 0, 1, "C")
            
            # Signature lines
            self.ln(10)
            col_width = 60
            
            # Left signature (Employee)
            self.cell(col_width, 4, "_________________________", 0, 0, "C")
            self.cell(30, 4, "", 0, 0, "C")
            
            # Right signature (Principal)
            self.cell(col_width, 4, "_________________________", 0, 1, "C")
            
            # Labels
            self.cell(col_width, 4, "Signature of Employee", 0, 0, "C")
            self.cell(30, 4, "", 0, 0, "C")
            self.cell(col_width, 4, "Principal III", 0, 1, "C")
            
            # Verification
            self.ln(5)
            self.cell(0, 4, "VERIFIED as to the prescribed office hours:", 0, 1, "C")
        
        def create_dtr_table(self, attendance_by_day, month, year, office_hours):
            """Create DTR table with exact template format"""
            days_in_month = calendar.monthrange(year, month)[1]
            
            # ========== TABLE HEADER ==========
            self.set_fill_color(240, 240, 240)  # Light gray for header
            self.set_y(140)
            
            # Column widths
            col_day = 12
            col_am_arrival = 14
            col_am_departure = 14
            col_pm_arrival = 14
            col_pm_departure = 14
            col_hours = 10
            col_minutes = 10
            
            # Total width
            total_width = col_day + col_am_arrival + col_am_departure + col_pm_arrival + col_pm_departure + col_hours + col_minutes
            
            # Center table
            x_start = (210 - total_width) / 2  # A4 width = 210mm
            self.set_x(x_start)
            
            # Day column
            self.set_font("Arial", "B", 10)
            self.cell(col_day, 12, "Day", 1, 0, "C", True)
            
            # A.M. (merged)
            self.cell(col_am_arrival + col_am_departure, 12, "A.M.", 1, 0, "C", True)
            
            # P.M. (merged)
            self.cell(col_pm_arrival + col_pm_departure, 12, "P.M.", 1, 0, "C", True)
            
            # Undertime (merged)
            self.cell(col_hours + col_minutes, 12, "Undertime", 1, 1, "C", True)
            
            # Sub-headers row
            self.set_x(x_start)
            self.set_font("Arial", "", 9)
            
            # Empty under Day
            self.cell(col_day, 8, "", 1, 0, "C")
            
            # AM sub-headers
            self.cell(col_am_arrival, 8, "Arrival", 1, 0, "C")
            self.cell(col_am_departure, 8, "Departure", 1, 0, "C")
            
            # PM sub-headers
            self.cell(col_pm_arrival, 8, "Arrival", 1, 0, "C")
            self.cell(col_pm_departure, 8, "Departure", 1, 0, "C")
            
            # Undertime sub-headers
            self.cell(col_hours, 8, "Hours", 1, 0, "C")
            self.cell(col_minutes, 8, "Minutes", 1, 1, "C")
            
            # ========== FILL DAYS ==========
            total_undertime_hours = 0
            total_undertime_minutes = 0
            
            for day in range(1, days_in_month + 1):
                date_obj = datetime(year, month, day)
                day_name = date_obj.strftime("%A").upper()
                
                self.set_x(x_start)
                
                # Day number (bold)
                self.set_font("Arial", "B", 10)
                self.cell(col_day, 8, str(day), 1, 0, "C")
                
                # Check for Saturday or Sunday
                if day_name == "SATURDAY":
                    self.set_font("Arial", "B", 10)
                    self.set_fill_color(240, 240, 240)
                    self.cell(col_am_arrival + col_am_departure, 8, "SATURDAY", 1, 0, "C", True)
                    self.set_fill_color(255, 255, 255)
                    
                    # Empty cells for PM
                    self.cell(col_pm_arrival, 8, "", 1, 0, "C")
                    self.cell(col_pm_departure, 8, "", 1, 0, "C")
                    
                    # Empty cells for undertime
                    self.cell(col_hours, 8, "", 1, 0, "C")
                    self.cell(col_minutes, 8, "", 1, 1, "C")
                    
                elif day_name == "SUNDAY":
                    self.set_font("Arial", "B", 10)
                    self.set_fill_color(240, 240, 240)
                    self.cell(col_am_arrival + col_am_departure, 8, "SUNDAY", 1, 0, "C", True)
                    self.set_fill_color(255, 255, 255)
                    
                    # Empty cells for PM
                    self.cell(col_pm_arrival, 8, "", 1, 0, "C")
                    self.cell(col_pm_departure, 8, "", 1, 0, "C")
                    
                    # Empty cells for undertime
                    self.cell(col_hours, 8, "", 1, 0, "C")
                    self.cell(col_minutes, 8, "", 1, 1, "C")
                    
                else:
                    # Regular work day
                    self.set_font("Arial", "", 10)
                    
                    if day in attendance_by_day and attendance_by_day[day]:
                        times = sorted(attendance_by_day[day])
                        
                        # AM times (before 12:00)
                        am_times = []
                        for t in times:
                            try:
                                hour = int(t.split(":")[0])
                                if hour < 12:
                                    am_times.append(t)
                            except:
                                continue
                        
                        # PM times (12:00 and after)
                        pm_times = []
                        for t in times:
                            try:
                                hour = int(t.split(":")[0])
                                if hour >= 12:
                                    pm_times.append(t)
                            except:
                                continue
                        
                        # Fill AM cells
                        if am_times:
                            self.cell(col_am_arrival, 8, am_times[0], 1, 0, "C")
                            self.cell(col_am_departure, 8, am_times[-1], 1, 0, "C")
                        else:
                            self.cell(col_am_arrival, 8, "", 1, 0, "C")
                            self.cell(col_am_departure, 8, "", 1, 0, "C")
                        
                        # Fill PM cells
                        if pm_times:
                            self.cell(col_pm_arrival, 8, pm_times[0], 1, 0, "C")
                            self.cell(col_pm_departure, 8, pm_times[-1], 1, 0, "C")
                        else:
                            self.cell(col_pm_arrival, 8, "", 1, 0, "C")
                            self.cell(col_pm_departure, 8, "", 1, 0, "C")
                        
                        # Calculate undertime
                        undertime_hours, undertime_minutes = self.calculate_undertime(
                            am_in=am_times[0] if am_times else None,
                            am_out=am_times[-1] if am_times else None,
                            pm_in=pm_times[0] if pm_times else None,
                            pm_out=pm_times[-1] if pm_times else None,
                            office_hours=office_hours
                        )
                        
                        self.cell(col_hours, 8, str(undertime_hours) if undertime_hours > 0 else "", 1, 0, "C")
                        self.cell(col_minutes, 8, str(undertime_minutes) if undertime_minutes > 0 else "", 1, 1, "C")
                        
                        total_undertime_hours += undertime_hours
                        total_undertime_minutes += undertime_minutes
                        
                    else:
                        # No data for this day
                        self.cell(col_am_arrival, 8, "", 1, 0, "C")
                        self.cell(col_am_departure, 8, "", 1, 0, "C")
                        self.cell(col_pm_arrival, 8, "", 1, 0, "C")
                        self.cell(col_pm_departure, 8, "", 1, 0, "C")
                        self.cell(col_hours, 8, "", 1, 0, "C")
                        self.cell(col_minutes, 8, "", 1, 1, "C")
            
            # ========== TOTAL ROW ==========
            self.set_x(x_start)
            self.set_font("Arial", "B", 10)
            self.cell(col_day + col_am_arrival + col_am_departure + col_pm_arrival + col_pm_departure, 
                    8, "TOTAL", 1, 0, "C")
            
            self.cell(col_hours, 8, str(total_undertime_hours) if total_undertime_hours > 0 else "", 1, 0, "C")
            self.cell(col_minutes, 8, str(total_undertime_minutes) if total_undertime_minutes > 0 else "", 1, 1, "C")
            
            # Update current Y position
            self.set_y(self.get_y() + 5)
        
        def calculate_undertime(self, am_in, am_out, pm_in, pm_out, office_hours):
            """Calculate undertime accurately"""
            if not am_in or not am_out or not pm_in or not pm_out:
                return 0, 0
            
            try:
                # Convert times to datetime
                def parse_time(t):
                    if isinstance(t, str):
                        return datetime.strptime(t, "%H:%M")
                    else:
                        return datetime.strptime(str(t), "%H:%M")
                
                am_in_time = parse_time(am_in)
                am_out_time = parse_time(am_out)
                pm_in_time = parse_time(pm_in)
                pm_out_time = parse_time(pm_out)
                
                # Office hours
                office_am_in = datetime.strptime(office_hours["regular_am_in"], "%H:%M")
                office_am_out = datetime.strptime(office_hours["regular_am_out"], "%H:%M")
                office_pm_in = datetime.strptime(office_hours["regular_pm_in"], "%H:%M")
                office_pm_out = datetime.strptime(office_hours["regular_pm_out"], "%H:%M")
                
                # Calculate expected time
                expected_total = (
                    (office_am_out - office_am_in).seconds / 60 +
                    (office_pm_out - office_pm_in).seconds / 60
                )
                
                # Calculate actual time
                actual_total = (
                    (am_out_time - am_in_time).seconds / 60 +
                    (pm_out_time - pm_in_time).seconds / 60
                )
                
                # Calculate undertime
                undertime_minutes = max(0, expected_total - actual_total)
                undertime_hours = int(undertime_minutes // 60)
                undertime_minutes_remainder = int(undertime_minutes % 60)
                
                return undertime_hours, undertime_minutes_remainder
                
            except:
                return 0, 0

    def generate_dtr_pdf(employee_no, employee_name, month, year, attendance_data, office_hours):
        """Generate DTR in PDF format following exact template"""
        
        # Process attendance data
        attendance_by_day = {}
        if attendance_data is not None and not attendance_data.empty:
            for _, row in attendance_data.iterrows():
                try:
                    day = int(row["Day"])
                    time_val = row["Time"]
                    
                    if hasattr(time_val, 'strftime'):
                        time_str = time_val.strftime("%H:%M")
                    elif isinstance(time_val, str):
                        time_str = time_val
                    else:
                        time_str = str(time_val)
                    
                    if day not in attendance_by_day:
                        attendance_by_day[day] = []
                    attendance_by_day[day].append(time_str)
                except:
                    continue
        
        # Sort times for each day
        for day in attendance_by_day:
            attendance_by_day[day] = sorted(attendance_by_day[day])
        
        # Create PDF
        pdf = DTR_PDF()
        
        # Format data
        month_name = calendar.month_name[month].upper()
        month_year = f"{month_name} {year}"
        
        office_hours_display = {
            'am': f"{office_hours['regular_am_in']} -- {office_hours['regular_am_out']}",
            'pm': f"{office_hours['regular_pm_in']} -- {office_hours['regular_pm_out']}",
            'saturday': office_hours['saturday']
        }
        
        # Add page and create DTR
        pdf.add_page()
        pdf.create_dtr_template(
            employee_no=employee_no,
            employee_name=employee_name,
            month_year=month_year,
            office_hours=office_hours_display,
            attendance_by_day=attendance_by_day,
            month_num=month,
            year_num=year,
            office_hours_dict=office_hours
        )
        
        # Add second DTR on same page (for printing front/back)
        pdf.add_page()
        pdf.create_dtr_template(
            employee_no=employee_no,
            employee_name=employee_name,
            month_year=month_year,
            office_hours=office_hours_display,
            attendance_by_day=attendance_by_day,
            month_num=month,
            year_num=year,
            office_hours_dict=office_hours
        )
        
        # Save to buffer
        buffer = BytesIO()
        pdf.output(buffer)
        buffer.seek(0)
        
        return buffer

except ImportError:
    st.error("Please install fpdf2: pip install fpdf2")
    st.stop()

# =============================================
# FILE PARSING FUNCTIONS
# =============================================

def parse_simple_attendance_file(uploaded_file):
    """Simple parser for attendance files"""
    try:
        content = uploaded_file.read().decode('utf-8', errors='ignore')
        lines = content.strip().split('\n')
        
        data = []
        for line_num, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            
            # Remove extra spaces
            line = ' '.join(line.split())
            
            # Split by space or tab
            if '\t' in line:
                parts = line.split('\t')
            else:
                parts = line.split()
            
            if len(parts) >= 3:
                emp_no = parts[0].strip()
                date_str = parts[1].strip()
                time_str = parts[2].strip()
                
                # Combine date and time
                datetime_str = f"{date_str} {time_str}"
                
                # Try different date formats
                date_formats = [
                    '%Y-%m-%d %H:%M:%S',
                    '%Y/%m/%d %H:%M:%S',
                    '%m/%d/%Y %H:%M:%S',
                    '%d/%m/%Y %H:%M:%S',
                    '%Y-%m-%d %H:%M',
                    '%m/%d/%Y %H:%M',
                    '%d-%m-%Y %H:%M:%S',
                    '%d-%m-%Y %H:%M',
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
                        'Time': dt.time().strftime('%H:%M'),
                        'Month': dt.month,
                        'Year': dt.year,
                        'Day': dt.day
                    })
        
        if data:
            return pd.DataFrame(data)
        else:
            # Try alternative parsing
            return parse_alternative_format(content)
            
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def parse_alternative_format(content):
    """Alternative parser for different formats"""
    lines = content.strip().split('\n')
    data = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # Look for patterns like: 7220970 2024-01-15 08:15:00
        pattern = r'(\d+)\s+(\d{4}[-/]\d{1,2}[-/]\d{1,2})\s+(\d{1,2}:\d{1,2}(?::\d{1,2})?)'
        match = re.search(pattern, line)
        
        if match:
            emp_no = match.group(1)
            date_str = match.group(2)
            time_str = match.group(3)
            
            # Standardize date format
            date_str = date_str.replace('/', '-')
            
            # Try to parse
            try:
                dt = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M:%S")
            except:
                try:
                    dt = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
                except:
                    continue
            
            data.append({
                'EmployeeNo': emp_no,
                'DateTime': dt,
                'Date': dt.date(),
                'Time': dt.time().strftime('%H:%M'),
                'Month': dt.month,
                'Year': dt.year,
                'Day': dt.day
            })
    
    if data:
        return pd.DataFrame(data)
    return None

def create_zip_file(pdf_files, month_name, year):
    """Create ZIP file containing all PDF files"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_info in pdf_files:
            clean_name = file_info["employee_name"].replace(",", "").replace(".", "").replace(" ", "_")
            filename = f"DTR_{clean_name}_{month_name}_{year}.pdf"
            zip_file.writestr(filename, file_info["pdf_file"].getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

# =============================================
# STREAMLIT APP - SIMPLE AND CLEAN
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

# Auto-load Richard P. Samoranos
if '7220970' not in st.session_state.employee_settings:
    st.session_state.employee_settings['7220970'] = {
        'name': 'SAMORANOS, RICHARD P.',
        'employee_no': '7220970'
    }

# Main Title
st.title("📋 DTR Generator - Civil Service Form No. 48")
st.markdown("**Manual National High School - Division of Davao del Sur**")
st.markdown("---")

# ========== STEP 1: FILE UPLOAD ==========
st.header("📤 1. Upload Attendance File")

uploaded_file = st.file_uploader(
    "Choose your .dat file",
    type=['dat', 'txt', 'csv'],
    help="Upload biometric attendance file"
)

if uploaded_file:
    with st.spinner("Reading file..."):
        df = parse_simple_attendance_file(uploaded_file)
        
        if df is not None and not df.empty:
            st.session_state.raw_data = df
            st.success(f"✅ Successfully loaded {len(df)} records")
            
            # Show preview
            with st.expander("👀 Preview Data"):
                st.dataframe(df.head(10))
        else:
            st.error("❌ Could not parse the file. Please check the format.")
            st.info("""
            **Expected format:**
            ```
            EmployeeID Date Time
            7220970 2024-01-15 08:15:00
            7220970 2024-01-15 12:00:00
            ```
            """)

# ========== STEP 2: OFFICE HOURS ==========
st.header("⏰ 2. Set Office Hours")

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

# ========== STEP 3: GENERATE DTR ==========
if st.session_state.raw_data is not None:
    df = st.session_state.raw_data
    
    st.header("🚀 3. Generate DTR")
    
    # Month selection
    if 'Month' in df.columns and 'Year' in df.columns:
        unique_months = df[['Month', 'Year']].drop_duplicates().sort_values(['Year', 'Month'])
        
        if not unique_months.empty:
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
                st.info(f"📊 Found {len(month_df)} records for {month_name} {year_num}")
                
                # Get employees
                employees_in_month = sorted(month_df['EmployeeNo'].unique())
                
                # Employee names
                st.subheader("✏️ Employee Names")
                
                for emp_id in employees_in_month:
                    current_name = st.session_state.employee_settings.get(
                        emp_id, 
                        {'name': f"EMPLOYEE {emp_id}"}
                    )['name']
                    
                    new_name = st.text_input(
                        f"Employee {emp_id}",
                        value=current_name,
                        key=f"name_{emp_id}"
                    )
                    
                    if new_name.strip():
                        st.session_state.employee_settings[emp_id] = {
                            'name': new_name.strip().upper(),
                            'employee_no': emp_id
                        }
                
                # Generate button
                st.markdown("---")
                
                if st.button("🚀 GENERATE DTR FILES", type="primary", use_container_width=True):
                    with st.spinner(f"Generating {len(employees_in_month)} DTR files..."):
                        pdf_files = []
                        
                        for emp_id in employees_in_month:
                            try:
                                emp_df = month_df[month_df['EmployeeNo'] == emp_id].copy()
                                
                                if emp_df.empty:
                                    continue
                                
                                emp_name = st.session_state.employee_settings.get(
                                    emp_id, 
                                    {'name': f"EMPLOYEE {emp_id}"}
                                )['name']
                                
                                # Generate PDF
                                pdf_file = generate_dtr_pdf(
                                    employee_no=emp_id,
                                    employee_name=emp_name,
                                    month=month_num,
                                    year=year_num,
                                    attendance_data=emp_df,
                                    office_hours=st.session_state.office_hours
                                )
                                
                                pdf_files.append({
                                    'employee_no': emp_id,
                                    'employee_name': emp_name,
                                    'pdf_file': pdf_file
                                })
                                
                            except Exception as e:
                                st.error(f"Error with {emp_id}: {str(e)}")
                        
                        if pdf_files:
                            st.success(f"✅ Generated {len(pdf_files)} DTR files!")
                            
                            # Create ZIP
                            zip_buffer = create_zip_file(pdf_files, month_name, year_num)
                            
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
                                with st.expander("📄 Individual Files"):
                                    for file_info in pdf_files:
                                        short_name = file_info['employee_name'][:20]
                                        st.download_button(
                                            label=f"⬇️ {short_name}",
                                            data=file_info['pdf_file'],
                                            file_name=f"DTR_{file_info['employee_name']}_{month_name}_{year_num}.pdf",
                                            mime="application/pdf",
                                            key=f"dl_{file_info['employee_no']}"
                                        )
                        else:
                            st.error("❌ No files were generated.")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Civil Service Form No. 48 DTR Generator | Version 6.0</p>
    <p><small>Manual National High School - Division of Davao del Sur</small></p>
</div>
""", unsafe_allow_html=True)

# Sample data button in sidebar
with st.sidebar:
    st.title("ℹ️ Information")
    
    if st.button("Load Sample Data"):
        # Create sample data for Richard P. Samoranos
        sample_data = []
        base_date = datetime(2024, 1, 1)
        
        for day in range(1, 16):  # First 15 days
            date = datetime(2024, 1, day)
            
            # Add morning and afternoon entries
            sample_data.append({
                'EmployeeNo': '7220970',
                'DateTime': datetime(2024, 1, day, 7, 30),
                'Date': date.date(),
                'Time': '07:30',
                'Month': 1,
                'Year': 2024,
                'Day': day
            })
            
            sample_data.append({
                'EmployeeNo': '7220970',
                'DateTime': datetime(2024, 1, day, 11, 50),
                'Date': date.date(),
                'Time': '11:50',
                'Month': 1,
                'Year': 2024,
                'Day': day
            })
            
            sample_data.append({
                'EmployeeNo': '7220970',
                'DateTime': datetime(2024, 1, day, 12, 50),
                'Date': date.date(),
                'Time': '12:50',
                'Month': 1,
                'Year': 2024,
                'Day': day
            })
            
            sample_data.append({
                'EmployeeNo': '7220970',
                'DateTime': datetime(2024, 1, day, 16, 30),
                'Date': date.date(),
                'Time': '16:30',
                'Month': 1,
                'Year': 2024,
                'Day': day
            })
        
        st.session_state.raw_data = pd.DataFrame(sample_data)
        st.success("✅ Sample data loaded! Refresh to see it.")
