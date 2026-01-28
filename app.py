import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar
from fpdf import FPDF
from io import BytesIO
import zipfile

# =============================================
# PDF DTR GENERATOR - FOLLOWING YOUR TEMPLATE
# =============================================

class DTR_PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
        self.add_font('Arial', '', 'c:/windows/fonts/arial.ttf', uni=True)
        self.add_font('Arial', 'B', 'c:/windows/fonts/arialbd.ttf', uni=True)
        self.add_font('Arial', 'I', 'c:/windows/fonts/ariali.ttf', uni=True)
    
    def header(self):
        # Header - Republic of the Philippines
        self.set_font('Arial', 'B', 11)
        self.cell(0, 5, "REPUBLIC OF THE PHILIPPINES", 0, 1, 'C')
        
        # Department of Education
        self.cell(0, 5, "Department of Education", 0, 1, 'C')
        
        # Division of Davao del Sur
        self.cell(0, 5, "Division of Davao del Sur", 0, 1, 'C')
        
        # Manual National High School
        self.cell(0, 5, "Manual National High School", 0, 1, 'C')
        
        self.ln(2)
        
        # Civil Service Form No. 48
        self.set_font('Arial', 'I', 9)
        self.cell(0, 4, "Civil Service Form No. 48", 0, 1, 'C')
        
        # Employee No.
        if hasattr(self, 'employee_no'):
            self.set_font('Arial', '', 9)
            self.cell(0, 4, f"Employee No. {self.employee_no}", 0, 1, 'C')
        
        self.ln(3)
        
        # DAILY TIME RECORD Title
        self.set_font('Arial', 'B', 14)
        self.cell(0, 8, "DAILY TIME RECORD", 0, 1, 'C')
        
        # Separator line
        self.set_font('Arial', '', 10)
        self.cell(0, 5, "-------------------------o0o-------------------------", 0, 1, 'C')
        
        self.ln(2)
        
        # Employee Name
        if hasattr(self, 'employee_name'):
            self.set_font('Arial', 'B', 12)
            self.cell(0, 6, self.employee_name, 0, 1, 'C')
            
            # (Name) label
            self.set_font('Arial', '', 9)
            self.cell(0, 4, "(Name)", 0, 1, 'C')
        
        self.ln(5)
        
        # Month and Year section
        if hasattr(self, 'month_year'):
            col_width = self.w / 3
            
            # Left side: Month and Year
            self.set_font('Arial', 'B', 11)
            self.cell(col_width, 6, self.month_year, 0, 0, 'C')
            
            # Middle space
            self.cell(col_width, 6, "", 0, 0, 'C')
            
            # Right side: "For the month of"
            self.set_font('Arial', 'I', 9)
            self.cell(col_width, 6, "For the month of", 0, 1, 'L')
            
            self.ln(3)
            
        # Office Hours section
        if hasattr(self, 'office_hours'):
            col_width = self.w / 3
            
            # AM Hours
            self.set_font('Arial', '', 10)
            self.cell(col_width, 6, self.office_hours['am'], 0, 0, 'C')
            
            # Middle: Official hours label
            self.set_font('Arial', 'I', 9)
            self.cell(col_width, 6, "Official hours for arrival and departure", 0, 0, 'L')
            
            # Right: Regular days label
            self.set_font('Arial', 'I', 9)
            self.cell(col_width, 6, "Regular days", 0, 1, 'L')
            
            self.ln(1)
            
            # PM Hours
            self.set_font('Arial', '', 10)
            self.cell(col_width, 6, self.office_hours['pm'], 0, 0, 'C')
            
            # Middle space
            self.set_font('Arial', '', 9)
            self.cell(col_width, 6, "", 0, 0, 'L')
            
            # Right: Saturdays label
            self.set_font('Arial', 'I', 9)
            self.cell(col_width, 6, "Saturdays", 0, 1, 'L')
            
            self.ln(1)
            
            # Saturday Hours
            self.set_font('Arial', '', 10)
            self.cell(col_width, 6, self.office_hours['saturday'], 0, 1, 'C')
        
        self.ln(5)
    
    def create_dtr_table(self, attendance_by_day, month, year, office_hours):
        """Create DTR table with data"""
        days_in_month = calendar.monthrange(year, month)[1]
        
        # Table header
        self.set_fill_color(240, 240, 240)
        self.set_font('Arial', 'B', 10)
        
        # Day column
        self.cell(10, 12, "Day", 1, 0, 'C', True)
        
        # AM Section (merged)
        self.cell(30, 12, "A.M.", 1, 0, 'C', True)
        
        # PM Section (merged)
        self.cell(30, 12, "P.M.", 1, 0, 'C', True)
        
        # Undertime Section (merged)
        self.cell(20, 12, "Undertime", 1, 1, 'C', True)
        
        # Subheaders row
        self.set_font('Arial', '', 9)
        
        # Day column empty
        self.cell(10, 8, "", 1, 0, 'C')
        
        # AM subheaders
        self.cell(15, 8, "Arrival", 1, 0, 'C')
        self.cell(15, 8, "Departure", 1, 0, 'C')
        
        # PM subheaders
        self.cell(15, 8, "Arrival", 1, 0, 'C')
        self.cell(15, 8, "Departure", 1, 0, 'C')
        
        # Undertime subheaders
        self.cell(10, 8, "Hours", 1, 0, 'C')
        self.cell(10, 8, "Minutes", 1, 1, 'C')
        
        self.set_font('Arial', '', 10)
        
        total_undertime_hours = 0
        total_undertime_minutes = 0
        
        # Fill days
        for day in range(1, days_in_month + 1):
            date_obj = datetime(year, month, day)
            day_name = date_obj.strftime('%A').upper()
            
            # Day number (bold)
            self.set_font('Arial', 'B', 10)
            self.cell(10, 8, str(day), 1, 0, 'C')
            self.set_font('Arial', '', 10)
            
            # Check if Saturday or Sunday
            if day_name == 'SATURDAY':
                self.set_font('Arial', 'B', 10)
                self.cell(30, 8, "SATURDAY", 1, 0, 'C')
                for _ in range(4):  # Empty cells for PM and Undertime
                    self.cell(15, 8, "", 1, 0, 'C')
                self.cell(20, 8, "", 1, 1, 'C')
                self.set_font('Arial', '', 10)
                
            elif day_name == 'SUNDAY':
                self.set_font('Arial', 'B', 10)
                self.cell(30, 8, "SUNDAY", 1, 0, 'C')
                for _ in range(4):  # Empty cells for PM and Undertime
                    self.cell(15, 8, "", 1, 0, 'C')
                self.cell(20, 8, "", 1, 1, 'C')
                self.set_font('Arial', '', 10)
                
            else:
                # Regular work day
                if day in attendance_by_day:
                    times = attendance_by_day[day]
                    
                    # AM times
                    am_times = [t for t in times if int(t.split(':')[0]) < 12]
                    if am_times:
                        self.cell(15, 8, am_times[0], 1, 0, 'C')
                        self.cell(15, 8, am_times[-1], 1, 0, 'C')
                    else:
                        self.cell(15, 8, "", 1, 0, 'C')
                        self.cell(15, 8, "", 1, 0, 'C')
                    
                    # PM times
                    pm_times = [t for t in times if int(t.split(':')[0]) >= 12]
                    if pm_times:
                        self.cell(15, 8, pm_times[0], 1, 0, 'C')
                        self.cell(15, 8, pm_times[-1], 1, 0, 'C')
                    else:
                        self.cell(15, 8, "", 1, 0, 'C')
                        self.cell(15, 8, "", 1, 0, 'C')
                    
                    # Calculate undertime
                    undertime_hours, undertime_minutes = self.calculate_undertime(
                        am_in=am_times[0] if am_times else None,
                        am_out=am_times[-1] if am_times else None,
                        pm_in=pm_times[0] if pm_times else None,
                        pm_out=pm_times[-1] if pm_times else None,
                        office_hours=office_hours
                    )
                    
                    self.cell(10, 8, str(undertime_hours) if undertime_hours else "", 1, 0, 'C')
                    self.cell(10, 8, str(undertime_minutes) if undertime_minutes else "", 1, 1, 'C')
                    
                    if undertime_hours:
                        total_undertime_hours += undertime_hours
                    if undertime_minutes:
                        total_undertime_minutes += undertime_minutes
                    
                else:
                    # No data for this day
                    for _ in range(4):  # Empty AM and PM cells
                        self.cell(15, 8, "", 1, 0, 'C')
                    self.cell(20, 8, "", 1, 1, 'C')
        
        # Total row
        self.set_font('Arial', 'B', 10)
        self.cell(70, 8, "TOTAL", 1, 0, 'C')
        self.cell(10, 8, str(total_undertime_hours) if total_undertime_hours else "", 1, 0, 'C')
        self.cell(10, 8, str(total_undertime_minutes) if total_undertime_minutes else "", 1, 1, 'C')
        
        self.ln(8)
        
        # Certification section
        self.set_font('Arial', '', 9)
        self.cell(0, 4, "I certify on my honor that the above is a true and correct report of the", 0, 1, 'C')
        self.cell(0, 4, "hours of work performed, record of which was made daily at the time of", 0, 1, 'C')
        self.cell(0, 4, "arrival and departure from office.", 0, 1, 'C')
        
        self.ln(8)
        
        # Signature lines
        col_width = self.w / 3
        
        # Left signature
        self.cell(col_width, 4, "_________________________", 0, 0, 'C')
        self.cell(col_width, 4, "", 0, 0, 'C')
        self.cell(col_width, 4, "_________________________", 0, 1, 'C')
        
        # Labels
        self.cell(col_width, 4, "Signature of Employee", 0, 0, 'C')
        self.cell(col_width, 4, "", 0, 0, 'C')
        self.cell(col_width, 4, "Principal III", 0, 1, 'C')
        
        self.ln(4)
        
        # Verification line
        self.cell(0, 4, "VERIFIED as to the prescribed office hours:", 0, 1, 'C')
    
    def calculate_undertime(self, am_in, am_out, pm_in, pm_out, office_hours):
        """Calculate undertime based on office hours"""
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
            total_expected = expected_am_hours + expected_pm_hours
            
            # Calculate actual hours
            actual_am_hours = (am_out_time - am_in_time).seconds / 3600 if am_out_time > am_in_time else 0
            actual_pm_hours = (pm_out_time - pm_in_time).seconds / 3600 if pm_out_time > pm_in_time else 0
            total_actual = actual_am_hours + actual_pm_hours
            
            # Calculate undertime
            undertime_decimal = max(0, total_expected - total_actual)
            undertime_hours = int(undertime_decimal)
            undertime_minutes = int((undertime_decimal - undertime_hours) * 60)
            
            return undertime_hours, undertime_minutes
        except:
            return 0, 0

def generate_dtr_pdf(employee_no, employee_name, month, year, attendance_data, office_hours):
    """Generate DTR in PDF format following the template"""
    
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
                else:
                    time_str = str(time_val)
                
                if day not in attendance_by_day:
                    attendance_by_day[day] = []
                attendance_by_day[day].append(time_str)
            except:
                continue
    
    # Create PDF
    pdf = DTR_PDF()
    pdf.employee_no = employee_no
    pdf.employee_name = employee_name.upper()
    
    # Format month and year
    month_name = calendar.month_name[month].upper()
    pdf.month_year = f"{month_name} {year}"
    
    # Format office hours for display
    pdf.office_hours = {
        'am': f"{office_hours['regular_am_in']} -- {office_hours['regular_am_out']}",
        'pm': f"{office_hours['regular_pm_in']} -- {office_hours['regular_pm_out']}",
        'saturday': office_hours['saturday']
    }
    
    # Add first page
    pdf.add_page()
    pdf.create_dtr_table(attendance_by_day, month, year, office_hours)
    
    # Add second page (duplicate for A4 paper - two DTRs per page)
    pdf.add_page()
    pdf.create_dtr_table(attendance_by_day, month, year, office_hours)
    
    # Save to buffer
    buffer = BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    
    return buffer

def create_zip_file(pdf_files, month_name, year):
    """Create ZIP file containing all PDF files"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_info in pdf_files:
            clean_name = file_info['employee_name'].replace(',', '').replace('.', '').replace(' ', '_')
            filename = f"DTR_{clean_name}_{month_name}_{year}.pdf"
            zip_file.writestr(filename, file_info['pdf_file'].getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

# =============================================
# STREAMLIT APP INTERFACE
# =============================================

# Page configuration
st.set_page_config(
    page_title="DTR Generator - Manual NHS",
    page_icon="📋",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        background-color: #1E3A8A;
        color: white;
        font-weight: bold;
    }
    .css-1d391kg {
        padding-top: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

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

# Main Title
st.markdown("""
    <h1 style='text-align: center; color: #1E3A8A; margin-bottom: 10px;'>
        📋 DTR Generator - Civil Service Form No. 48
    </h1>
""", unsafe_allow_html=True)

st.markdown("""
    <p style='text-align: center; color: #4B5563; margin-bottom: 30px;'>
        Manual National High School - Division of Davao del Sur
    </p>
    <hr>
""", unsafe_allow_html=True)

# ========== STEP 1: FILE UPLOAD ==========
st.header("1️⃣ Upload Biometric Attendance File")
uploaded_file = st.file_uploader(
    "Choose your attendance file (.dat, .txt, .csv)",
    type=['dat', 'txt', 'csv'],
    help="Upload the file exported from your biometric system"
)

if uploaded_file:
    try:
        # Read and parse the file
        content = uploaded_file.read().decode('utf-8', errors='ignore')
        lines = [line.strip() for line in content.split('\n') if line.strip()]
        
        data = []
        for line in lines:
            parts = line.split()
            if len(parts) >= 2:
                emp_no = parts[0].strip()
                datetime_str = ' '.join(parts[1:3]) if len(parts) >= 3 else parts[1]
                
                # Try to parse datetime
                for fmt in ['%Y-%m-%d %H:%M:%S', '%m/%d/%Y %H:%M:%S', '%d/%m/%Y %H:%M:%S']:
                    try:
                        dt = datetime.strptime(datetime_str, fmt)
                        break
                    except:
                        dt = None
                
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
            df = pd.DataFrame(data)
            st.session_state.raw_data = df
            
            st.success(f"✅ Successfully loaded {len(df)} attendance records")
            
            # Show statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", len(df))
            with col2:
                st.metric("Unique Employees", df['EmployeeNo'].nunique())
            with col3:
                date_min = df['Date'].min().strftime('%m/%d/%Y') if hasattr(df['Date'].min(), 'strftime') else df['Date'].min()
                date_max = df['Date'].max().strftime('%m/%d/%Y') if hasattr(df['Date'].max(), 'strftime') else df['Date'].max()
                st.metric("Date Range", f"{date_min} to {date_max}")
            
            # Preview data
            with st.expander("👀 Preview Attendance Data"):
                st.dataframe(df.head(20))
        else:
            st.error("❌ No valid attendance records found in the file.")
            
    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

# ========== STEP 2: OFFICE HOURS SETTING ==========
st.header("2️⃣ Set Office Hours")

st.info("⚠️ **IMPORTANT:** Set the official office hours for regular work days.")

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

# Save to session state
st.session_state.office_hours = {
    'regular_am_in': am_in,
    'regular_am_out': am_out,
    'regular_pm_in': pm_in,
    'regular_pm_out': pm_out,
    'saturday': saturday_hours
}

# ========== STEP 3: PROCESS DATA ==========
if st.session_state.raw_data is not None:
    df = st.session_state.raw_data
    
    st.header("3️⃣ Generate DTR Files")
    
    # Month selection
    if 'Month' in df.columns and 'Year' in df.columns:
        unique_months = df[['Month', 'Year']].drop_duplicates().sort_values(['Year', 'Month'])
        
        if not unique_months.empty:
            # Create month options
            month_options = []
            for _, row in unique_months.iterrows():
                month_name = calendar.month_name[row['Month']]
                month_options.append(f"{month_name} {row['Year']}")
            
            selected_month = st.selectbox("Select Month for DTR Generation", month_options)
            
            # Parse selection
            month_name, year_str = selected_month.split()
            month_num = list(calendar.month_name).index(month_name)
            year_num = int(year_str)
            
            # Filter data for selected month
            month_df = df[(df['Month'] == month_num) & (df['Year'] == year_num)].copy()
            
            if not month_df.empty:
                st.info(f"📊 **Found {len(month_df)} attendance records for {month_name} {year_num}**")
                
                # Get employees in this month
                employees_in_month = sorted(month_df['EmployeeNo'].unique())
                
                # PRE-LOAD Richard P. Samoranos
                if '7220970' not in st.session_state.employee_settings:
                    st.session_state.employee_settings['7220970'] = {
                        'name': 'SAMORANOS, RICHARD P.',
                        'employee_no': '7220970'
                    }
                
                # Employee names management
                st.subheader("✏️ Employee Names Management")
                
                # Create a form for editing names
                with st.form(key='employee_names_form'):
                    st.write("**Edit employee names for each biometric ID:**")
                    
                    # Container for employee name inputs
                    names_container = st.container()
                    
                    employee_names = {}
                    with names_container:
                        cols = st.columns(2)
                        emp_count = len(employees_in_month)
                        
                        for idx, emp_id in enumerate(employees_in_month):
                            col_idx = idx % 2
                            with cols[col_idx]:
                                # Get current name
                                current_settings = st.session_state.employee_settings.get(
                                    emp_id, 
                                    {'name': f"EMPLOYEE {emp_id}", 'employee_no': emp_id}
                                )
                                current_name = current_settings['name']
                                
                                # Create input field
                                new_name = st.text_input(
                                    f"**ID: {emp_id}**",
                                    value=current_name,
                                    key=f"name_{emp_id}",
                                    help=f"Enter name for employee with biometric ID {emp_id}"
                                )
                                
                                employee_names[emp_id] = new_name
                    
                    # Submit button for names
                    if st.form_submit_button("💾 Save All Names"):
                        for emp_id, emp_name in employee_names.items():
                            if emp_name.strip():
                                st.session_state.employee_settings[emp_id] = {
                                    'name': emp_name.strip().upper(),
                                    'employee_no': emp_id
                                }
                        st.success("✅ All employee names saved successfully!")
                
                st.markdown("---")
                
                # Generate DTR Button
                if st.button("🚀 GENERATE PDF DTR FILES NOW", type="primary", use_container_width=True):
                    with st.spinner(f"Generating DTR PDF files for {len(employees_in_month)} employees..."):
                        pdf_files = []
                        success_count = 0
                        errors = []
                        
                        # Progress bar
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        for idx, emp_id in enumerate(employees_in_month):
                            status_text.text(f"Processing: Employee {emp_id} ({idx+1}/{len(employees_in_month)})")
                            
                            try:
                                # Filter employee data
                                emp_df = month_df[month_df['EmployeeNo'] == emp_id].copy()
                                
                                if emp_df.empty:
                                    errors.append(f"❌ {emp_id}: No attendance data found")
                                    continue
                                
                                # Get employee name
                                emp_settings = st.session_state.employee_settings.get(
                                    emp_id, 
                                    {'name': f"EMPLOYEE {emp_id}", 'employee_no': emp_id}
                                )
                                emp_name = emp_settings['name']
                                
                                # Generate PDF DTR
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
                                
                                success_count += 1
                                
                            except Exception as e:
                                errors.append(f"❌ {emp_id}: {str(e)}")
                            
                            # Update progress
                            progress_bar.progress((idx + 1) / len(employees_in_month))
                        
                        # Clear progress indicators
                        progress_bar.empty()
                        status_text.empty()
                        
                        # Show results
                        if pdf_files:
                            st.success(f"✅ Successfully generated {success_count} DTR PDF files!")
                            
                            # Show Richard P. Samoranos special note
                            richard_pdf = next((f for f in pdf_files if f['employee_no'] == '7220970'), None)
                            if richard_pdf:
                                st.info(f"📄 **Special Note:** DTR for Richard P. Samoranos (Employee No. 7220970) has been generated successfully.")
                            
                            # Create ZIP file
                            zip_buffer = create_zip_file(pdf_files, month_name, year_num)
                            
                            # Download buttons
                            st.markdown("### 📥 Download Options")
                            
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                st.download_button(
                                    label="📦 DOWNLOAD ALL FILES (ZIP)",
                                    data=zip_buffer,
                                    file_name=f"DTR_{month_name}_{year_num}.zip",
                                    mime="application/zip",
                                    use_container_width=True,
                                    help="Download all DTR PDF files in a single ZIP archive"
                                )
                            
                            with col2:
                                with st.expander("📄 Download Individual PDF Files"):
                                    for file_info in pdf_files:
                                        # Format display name
                                        display_name = file_info['employee_name']
                                        if len(display_name) > 25:
                                            display_name = display_name[:22] + "..."
                                        
                                        # Special styling for Richard
                                        if file_info['employee_no'] == '7220970':
                                            st.success(f"✅ **{display_name}** (ID: {file_info['employee_no']})")
                                        else:
                                            st.write(f"{display_name} (ID: {file_info['employee_no']})")
                                        
                                        # Download button
                                        st.download_button(
                                            label=f"⬇️ Download {display_name}",
                                            data=file_info['pdf_file'],
                                            file_name=f"DTR_{file_info['employee_name']}_{month_name}_{year_num}.pdf",
                                            mime="application/pdf",
                                            key=f"dl_{file_info['employee_no']}"
                                        )
                            
                            # Show errors if any
                            if errors:
                                with st.expander("⚠️ View Processing Errors"):
                                    for error in errors:
                                        st.write(error)
                        else:
                            st.error("❌ No DTR files were generated.")
                            
                            if errors:
                                with st.expander("Error Details"):
                                    st.write("**Possible issues:**")
                                    for error in errors:
                                        st.write(f"- {error}")
            else:
                st.warning(f"⚠️ No attendance data found for {month_name} {year_num}")

# Footer
st.markdown("---")
st.markdown("""
    <div style="text-align: center; color: #6B7280; padding: 20px;">
        <p><strong>Civil Service Form No. 48 - DTR Generator</strong><br>
        <small>Version 2.0 | Manual National High School | Division of Davao del Sur</small></p>
    </div>
""", unsafe_allow_html=True)

# Installation instructions in sidebar
with st.sidebar:
    st.title("📋 About This App")
    
    st.markdown("""
    ### Features:
    ✅ **PDF Output** - Generates DTR in PDF format
    ✅ **Two DTRs per A4** - Following Civil Service format
    ✅ **Pre-loaded Names** - Richard P. Samoranos (7220970)
    ✅ **Batch Processing** - Generate all employees at once
    ✅ **ZIP Download** - All files in one archive
    
    ### Installation Requirements:
    ```bash
    pip install streamlit pandas fpdf
    ```
    
    ### How to Use:
    1. Upload biometric attendance file
    2. Set office hours
    3. Edit employee names (if needed)
    4. Select month
    5. Generate and download PDFs
    
    ### File Format:
    ```
    EmployeeNo Date Time
    7220970 2026-01-05 06:51:00
    7220970 2026-01-05 11:54:00
    ```
    """)
