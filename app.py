import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import calendar
import zipfile
from io import BytesIO
import math

# =============================================
# PDF DTR GENERATOR - PERFECT MATCH TO TEMPLATE
# =============================================

try:
    from fpdf import FPDF
    
    class DTR_PDF(FPDF):
        def __init__(self):
            super().__init__()
            self.set_auto_page_break(auto=True, margin=15)
            
        def header(self):
            # Set font for the entire header
            self.set_font("Arial", "B", 11)
            
            # First header section - TOP LEFT
            self.set_y(10)
            self.cell(0, 5, "REPUBLIC OF THE PHILIPPINES", 0, 1, "L")
            self.set_font("Arial", "B", 10)
            self.cell(0, 5, "Department of Education", 0, 1, "L")
            self.cell(0, 5, "Division of Davao del Sur", 0, 1, "L")
            self.cell(0, 5, "Manual National High School", 0, 1, "L")
            
            # Civil Service Form No. 48 - CENTER
            self.set_y(10)
            self.set_font("Arial", "I", 9)
            self.cell(0, 5, "Civil Service Form No. 48", 0, 0, "C")
            
            # Employee No. - CENTER (below)
            self.ln(5)
            if hasattr(self, 'employee_no'):
                self.set_font("Arial", "", 9)
                self.cell(0, 5, f"Employee No. {self.employee_no}", 0, 1, "C")
            
            # DAILY TIME RECORD Title - CENTER
            self.ln(5)
            self.set_font("Arial", "B", 14)
            self.cell(0, 8, "DAILY TIME RECORD", 0, 1, "C")
            
            # Separator line
            self.set_font("Arial", "", 10)
            self.cell(0, 5, "-------------------------o0o-------------------------", 0, 1, "C")
            
            # Employee Name - CENTER
            self.ln(3)
            if hasattr(self, 'employee_name'):
                self.set_font("Arial", "B", 12)
                self.cell(0, 6, self.employee_name, 0, 1, "C")
                
                # (Name) label
                self.set_font("Arial", "", 9)
                self.cell(0, 4, "(Name)", 0, 1, "C")
            
            self.ln(5)
            
            # ========== MONTH AND OFFICE HOURS SECTION ==========
            # Month and Year - LEFT SIDE
            if hasattr(self, 'month_year'):
                self.set_font("Arial", "B", 11)
                self.cell(50, 6, self.month_year, 0, 0, "L")
            
            # Office hours - LEFT SIDE (below month)
            if hasattr(self, 'office_hours'):
                self.ln(6)
                self.set_font("Arial", "", 10)
                
                # AM Hours
                self.cell(50, 5, self.office_hours['am'], 0, 0, "L")
                self.ln(5)
                
                # PM Hours
                self.cell(50, 5, self.office_hours['pm'], 0, 0, "L")
                self.ln(5)
                
                # Saturday Hours
                self.cell(50, 5, self.office_hours['saturday'], 0, 0, "L")
            
            # Right side labels
            self.set_y(65)  # Position for right side labels
            self.set_x(140)  # Right side position
            
            # "For the month of" label
            self.set_font("Arial", "I", 9)
            self.cell(50, 5, "For the month of", 0, 2, "L")
            
            self.ln(3)
            self.set_x(140)
            self.cell(50, 5, "Official hours for arrival", 0, 2, "L")
            self.set_x(140)
            self.cell(50, 5, "and departure", 0, 2, "L")
            
            self.ln(3)
            self.set_x(140)
            self.set_font("Arial", "I", 9)
            self.cell(30, 5, "Regular days", 0, 2, "L")
            
            self.set_x(140)
            self.cell(30, 5, "Saturday", 0, 2, "L")
            
            self.ln(10)
        
        def create_dtr_table(self, attendance_by_day, month, year, office_hours):
            """Create DTR table with EXACT template format"""
            days_in_month = calendar.monthrange(year, month)[1]
            
            # ========== TABLE HEADER ==========
            self.set_fill_color(220, 220, 220)  # Light gray for header
            
            # Day column header
            self.set_font("Arial", "B", 10)
            self.cell(12, 12, "Day", 1, 0, "C", True)
            
            # A.M. header (merged)
            self.cell(28, 12, "A.M.", 1, 0, "C", True)
            
            # P.M. header (merged)
            self.cell(28, 12, "P.M.", 1, 0, "C", True)
            
            # Undertime header (merged)
            self.cell(20, 12, "Undertime", 1, 1, "C", True)
            
            # ========== SUB-HEADERS ==========
            # Empty cell under "Day"
            self.set_font("Arial", "", 9)
            self.cell(12, 8, "", 1, 0, "C")
            
            # A.M. sub-headers
            self.cell(14, 8, "Arrival", 1, 0, "C")
            self.cell(14, 8, "Departure", 1, 0, "C")
            
            # P.M. sub-headers
            self.cell(14, 8, "Arrival", 1, 0, "C")
            self.cell(14, 8, "Departure", 1, 0, "C")
            
            # Undertime sub-headers
            self.cell(10, 8, "Hours", 1, 0, "C")
            self.cell(10, 8, "Minutes", 1, 1, "C")
            
            # ========== FILL DAYS ==========
            total_undertime_hours = 0
            total_undertime_minutes = 0
            
            for day in range(1, days_in_month + 1):
                date_obj = datetime(year, month, day)
                day_name = date_obj.strftime("%A").upper()
                
                # Day cell (always bold)
                self.set_font("Arial", "B", 10)
                self.cell(12, 8, str(day), 1, 0, "C")
                self.set_font("Arial", "", 10)
                
                # Check for SATURDAY or SUNDAY
                if day_name == "SATURDAY":
                    self.set_font("Arial", "B", 10)
                    self.set_fill_color(240, 240, 240)  # Light gray fill
                    self.cell(28, 8, "SATURDAY", 1, 0, "C", True)
                    self.set_fill_color(255, 255, 255)  # Reset fill
                    
                    # Empty cells for PM columns
                    self.cell(14, 8, "", 1, 0, "C")
                    self.cell(14, 8, "", 1, 0, "C")
                    
                    # Empty cells for Undertime
                    self.cell(10, 8, "", 1, 0, "C")
                    self.cell(10, 8, "", 1, 1, "C")
                    
                elif day_name == "SUNDAY":
                    self.set_font("Arial", "B", 10)
                    self.set_fill_color(240, 240, 240)  # Light gray fill
                    self.cell(28, 8, "SUNDAY", 1, 0, "C", True)
                    self.set_fill_color(255, 255, 255)  # Reset fill
                    
                    # Empty cells for PM columns
                    self.cell(14, 8, "", 1, 0, "C")
                    self.cell(14, 8, "", 1, 0, "C")
                    
                    # Empty cells for Undertime
                    self.cell(10, 8, "", 1, 0, "C")
                    self.cell(10, 8, "", 1, 1, "C")
                    
                else:
                    # REGULAR WORK DAY
                    if day in attendance_by_day and attendance_by_day[day]:
                        times = sorted(attendance_by_day[day])
                        
                        # Get AM times (before 12:00)
                        am_times = []
                        for t in times:
                            try:
                                if isinstance(t, str):
                                    hour = int(t.split(":")[0])
                                else:
                                    hour = t.hour
                                if hour < 12:
                                    am_times.append(t)
                            except:
                                continue
                        
                        # Get PM times (12:00 and after)
                        pm_times = []
                        for t in times:
                            try:
                                if isinstance(t, str):
                                    hour = int(t.split(":")[0])
                                else:
                                    hour = t.hour
                                if hour >= 12:
                                    pm_times.append(t)
                            except:
                                continue
                        
                        # Format AM times
                        if am_times:
                            am_in = am_times[0] if isinstance(am_times[0], str) else am_times[0].strftime("%H:%M")
                            am_out = am_times[-1] if isinstance(am_times[-1], str) else am_times[-1].strftime("%H:%M")
                            self.cell(14, 8, am_in, 1, 0, "C")
                            self.cell(14, 8, am_out, 1, 0, "C")
                        else:
                            self.cell(14, 8, "", 1, 0, "C")
                            self.cell(14, 8, "", 1, 0, "C")
                        
                        # Format PM times
                        if pm_times:
                            pm_in = pm_times[0] if isinstance(pm_times[0], str) else pm_times[0].strftime("%H:%M")
                            pm_out = pm_times[-1] if isinstance(pm_times[-1], str) else pm_times[-1].strftime("%H:%M")
                            self.cell(14, 8, pm_in, 1, 0, "C")
                            self.cell(14, 8, pm_out, 1, 0, "C")
                        else:
                            self.cell(14, 8, "", 1, 0, "C")
                            self.cell(14, 8, "", 1, 0, "C")
                        
                        # Calculate undertime
                        undertime_hours, undertime_minutes = self.calculate_undertime(
                            am_in=am_times[0] if am_times else None,
                            am_out=am_times[-1] if am_times else None,
                            pm_in=pm_times[0] if pm_times else None,
                            pm_out=pm_times[-1] if pm_times else None,
                            office_hours=office_hours
                        )
                        
                        # Display undertime
                        self.cell(10, 8, str(undertime_hours) if undertime_hours > 0 else "", 1, 0, "C")
                        self.cell(10, 8, str(undertime_minutes) if undertime_minutes > 0 else "", 1, 1, "C")
                        
                        # Add to totals
                        total_undertime_hours += undertime_hours
                        total_undertime_minutes += undertime_minutes
                        
                    else:
                        # No attendance data for this day
                        for _ in range(4):  # Empty AM/PM cells
                            self.cell(14, 8, "", 1, 0, "C")
                        self.cell(20, 8, "", 1, 1, "C")
            
            # ========== TOTAL ROW ==========
            self.set_font("Arial", "B", 10)
            self.cell(68, 8, "TOTAL", 1, 0, "C")  # Span first 4 columns
            self.cell(10, 8, str(total_undertime_hours) if total_undertime_hours > 0 else "", 1, 0, "C")
            self.cell(10, 8, str(total_undertime_minutes) if total_undertime_minutes > 0 else "", 1, 1, "C")
            
            self.ln(10)
            
            # ========== CERTIFICATION SECTION ==========
            self.set_font("Arial", "", 9)
            self.cell(0, 4, "I certify on my honor that the above is a true and correct report of the", 0, 1, "C")
            self.cell(0, 4, "hours of work performed, record of which was made daily at the time of", 0, 1, "C")
            self.cell(0, 4, "arrival and departure from office.", 0, 1, "C")
            
            self.ln(8)
            
            # Signature lines
            col_width = 60
            
            # Left signature line (Employee)
            self.cell(col_width, 4, "_________________________", 0, 0, "C")
            self.cell(30, 4, "", 0, 0, "C")  # Spacing
            # Right signature line (Principal)
            self.cell(col_width, 4, "_________________________", 0, 1, "C")
            
            # Labels under signatures
            self.cell(col_width, 4, "Signature of Employee", 0, 0, "C")
            self.cell(30, 4, "", 0, 0, "C")  # Spacing
            self.cell(col_width, 4, "Principal III", 0, 1, "C")
            
            self.ln(5)
            
            # Verification line
            self.set_font("Arial", "", 9)
            self.cell(0, 4, "VERIFIED as to the prescribed office hours:", 0, 1, "C")
        
        def calculate_undertime(self, am_in, am_out, pm_in, pm_out, office_hours):
            """Calculate undertime accurately"""
            if not am_in or not am_out or not pm_in or not pm_out:
                return 0, 0
            
            try:
                # Convert times to datetime objects
                def parse_time(t):
                    if isinstance(t, str):
                        return datetime.strptime(t, "%H:%M")
                    elif hasattr(t, 'strftime'):
                        return datetime.combine(datetime.today(), t)
                    else:
                        return datetime.strptime(str(t), "%H:%M")
                
                am_in_time = parse_time(am_in)
                am_out_time = parse_time(am_out)
                pm_in_time = parse_time(pm_in)
                pm_out_time = parse_time(pm_out)
                
                # Parse office hours
                office_am_in = datetime.strptime(office_hours["regular_am_in"], "%H:%M")
                office_am_out = datetime.strptime(office_hours["regular_am_out"], "%H:%M")
                office_pm_in = datetime.strptime(office_hours["regular_pm_in"], "%H:%M")
                office_pm_out = datetime.strptime(office_hours["regular_pm_out"], "%H:%M")
                
                # Calculate expected total minutes
                expected_total_minutes = (
                    (office_am_out - office_am_in).seconds / 60 +
                    (office_pm_out - office_pm_in).seconds / 60
                )
                
                # Calculate actual total minutes
                actual_total_minutes = (
                    (am_out_time - am_in_time).seconds / 60 +
                    (pm_out_time - pm_in_time).seconds / 60
                )
                
                # Calculate undertime in minutes
                undertime_minutes = max(0, expected_total_minutes - actual_total_minutes)
                
                # Convert to hours and minutes
                undertime_hours = int(undertime_minutes // 60)
                undertime_minutes_remainder = int(undertime_minutes % 60)
                
                return undertime_hours, undertime_minutes_remainder
                
            except Exception as e:
                print(f"Error calculating undertime: {e}")
                return 0, 0
    
    def generate_dtr_pdf(employee_no, employee_name, month, year, attendance_data, office_hours):
        """Generate DTR in PDF format following exact template"""
        
        # Process attendance data by day
        attendance_by_day = {}
        if attendance_data is not None and not attendance_data.empty:
            for _, row in attendance_data.iterrows():
                try:
                    day = int(row["Day"])
                    
                    # Get time value
                    time_val = row["Time"]
                    
                    # Convert to string format
                    if hasattr(time_val, 'strftime'):
                        time_str = time_val.strftime("%H:%M")
                    elif isinstance(time_val, str):
                        time_str = time_val
                    else:
                        time_str = str(time_val)
                    
                    # Ensure time is in HH:MM format
                    if ":" in time_str:
                        if day not in attendance_by_day:
                            attendance_by_day[day] = []
                        attendance_by_day[day].append(time_str)
                        
                except Exception as e:
                    continue
        
        # Sort times for each day
        for day in attendance_by_day:
            attendance_by_day[day] = sorted(attendance_by_day[day])
        
        # Create PDF
        pdf = DTR_PDF()
        pdf.employee_no = employee_no
        pdf.employee_name = employee_name.upper()
        
        # Format month and year
        month_name = calendar.month_name[month].upper()
        pdf.month_year = f"{month_name} {year}"
        
        # Format office hours
        pdf.office_hours = {
            'am': f"{office_hours['regular_am_in']} -- {office_hours['regular_am_out']}",
            'pm': f"{office_hours['regular_pm_in']} -- {office_hours['regular_pm_out']}",
            'saturday': office_hours['saturday']
        }
        
        # Add page and create table
        pdf.add_page()
        pdf.create_dtr_table(attendance_by_day, month, year, office_hours)
        
        # Save to buffer
        buffer = BytesIO()
        pdf.output(buffer)
        buffer.seek(0)
        
        return buffer
        
except ImportError:
    st.error("""
    ## ❌ Missing Dependencies
    
    Please install the required package:
    
    **For local development:**
    ```
    pip install fpdf2
    ```
    
    **For Streamlit Cloud:**
    Create a `requirements.txt` file with:
    ```
    streamlit>=1.28.0
    pandas>=2.0.0
    numpy>=1.24.0
    fpdf2>=2.7.4
    ```
    """)
    st.stop()

# =============================================
# HELPER FUNCTIONS
# =============================================

def create_zip_file(pdf_files, month_name, year):
    """Create ZIP file containing all PDF files"""
    zip_buffer = BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for file_info in pdf_files:
            # Clean filename
            clean_name = file_info["employee_name"].replace(",", "").replace(".", "").replace(" ", "_")
            filename = f"DTR_{clean_name}_{month_name}_{year}.pdf"
            zip_file.writestr(filename, file_info["pdf_file"].getvalue())
    
    zip_buffer.seek(0)
    return zip_buffer

def parse_attendance_file(uploaded_file):
    """Parse various attendance file formats"""
    try:
        # Read file content
        content = uploaded_file.read().decode("utf-8", errors="ignore")
        lines = [line.strip() for line in content.split("\n") if line.strip()]
        
        data = []
        for line in lines:
            # Try different delimiters
            if "\t" in line:
                parts = line.split("\t")
            else:
                parts = line.split()
            
            if len(parts) >= 2:
                emp_no = parts[0].strip()
                datetime_str = " ".join(parts[1:3]) if len(parts) >= 3 else parts[1]
                
                # Try different datetime formats
                dt = None
                date_formats = [
                    "%Y-%m-%d %H:%M:%S",
                    "%m/%d/%Y %H:%M:%S", 
                    "%d/%m/%Y %H:%M:%S",
                    "%Y/%m/%d %H:%M:%S",
                    "%m-%d-%Y %H:%M:%S",
                    "%d-%m-%Y %H:%M:%S"
                ]
                
                for fmt in date_formats:
                    try:
                        dt = datetime.strptime(datetime_str, fmt)
                        break
                    except:
                        continue
                
                if dt:
                    data.append({
                        "EmployeeNo": emp_no,
                        "DateTime": dt,
                        "Date": dt.date(),
                        "Time": dt.time().strftime("%H:%M"),
                        "Month": dt.month,
                        "Year": dt.year,
                        "Day": dt.day
                    })
        
        if data:
            return pd.DataFrame(data)
        else:
            return None
            
    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return None

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
        padding: 0.5rem 1rem;
    }
    
    .stButton>button:hover {
        background-color: #2D4A9A;
    }
    
    .success-box {
        background-color: #D4EDDA;
        color: #155724;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #28A745;
        margin: 1rem 0;
    }
    
    .warning-box {
        background-color: #FFF3CD;
        color: #856404;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #FFC107;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if "raw_data" not in st.session_state:
    st.session_state.raw_data = None
if "employee_settings" not in st.session_state:
    st.session_state.employee_settings = {}
if "office_hours" not in st.session_state:
    st.session_state.office_hours = {
        "regular_am_in": "07:30",
        "regular_am_out": "11:50",
        "regular_pm_in": "12:50",
        "regular_pm_out": "16:30",
        "saturday": "AS REQUIRED"
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
""", unsafe_allow_html=True)

# ========== STEP 1: FILE UPLOAD ==========
with st.container():
    st.header("1️⃣ Upload Biometric Attendance File")
    
    uploaded_file = st.file_uploader(
        "Choose your attendance file (.dat, .txt, .csv)",
        type=["dat", "txt", "csv"],
        help="Upload the file exported from your biometric system"
    )
    
    if uploaded_file:
        with st.spinner("Processing attendance file..."):
            df = parse_attendance_file(uploaded_file)
            
            if df is not None and not df.empty:
                st.session_state.raw_data = df
                
                st.markdown(f"""
                <div class="success-box">
                    <strong>✅ Successfully loaded {len(df)} attendance records</strong>
                </div>
                """, unsafe_allow_html=True)
                
                # Show statistics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Records", len(df))
                with col2:
                    st.metric("Unique Employees", df["EmployeeNo"].nunique())
                with col3:
                    months = df["Month"].unique()
                    st.metric("Months Covered", len(months))
                with col4:
                    years = df["Year"].unique()
                    st.metric("Years Covered", len(years))
                
                # Preview data
                with st.expander("👀 Preview Attendance Data"):
                    st.dataframe(df.head(20))
            else:
                st.error("❌ No valid attendance records found in the file.")

# ========== STEP 2: OFFICE HOURS SETTING ==========
with st.container():
    st.header("2️⃣ Set Office Hours")
    
    st.info("⚠️ **Set the official office hours for regular work days**")
    
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
        "regular_am_in": am_in,
        "regular_am_out": am_out,
        "regular_pm_in": pm_in,
        "regular_pm_out": pm_out,
        "saturday": saturday_hours
    }

# ========== STEP 3: PROCESS DATA ==========
if st.session_state.raw_data is not None:
    df = st.session_state.raw_data
    
    with st.container():
        st.header("3️⃣ Generate DTR Files")
        
        # Month selection
        if "Month" in df.columns and "Year" in df.columns:
            # Get unique months
            unique_months = df[["Month", "Year"]].drop_duplicates().sort_values(["Year", "Month"])
            
            if not unique_months.empty:
                # Create month options
                month_options = []
                for _, row in unique_months.iterrows():
                    month_name = calendar.month_name[row["Month"]]
                    month_options.append(f"{month_name} {row['Year']}")
                
                # Add "All Months" option
                month_options.insert(0, "All Months")
                
                selected_month = st.selectbox("Select Month for DTR Generation", month_options)
                
                if selected_month != "All Months":
                    # Parse selection for single month
                    month_name, year_str = selected_month.split()
                    month_num = list(calendar.month_name).index(month_name)
                    year_num = int(year_str)
                    
                    # Filter data for selected month
                    month_df = df[(df["Month"] == month_num) & (df["Year"] == year_num)].copy()
                    
                    if not month_df.empty:
                        st.markdown(f"""
                        <div class="success-box">
                            <strong>📊 Found {len(month_df)} attendance records for {month_name} {year_num}</strong>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        # Get employees in this month
                        employees_in_month = sorted(month_df["EmployeeNo"].unique())
                        
                        # PRE-LOAD Richard P. Samoranos
                        if "7220970" not in st.session_state.employee_settings:
                            st.session_state.employee_settings["7220970"] = {
                                "name": "SAMORANOS, RICHARD P.",
                                "employee_no": "7220970"
                            }
                        
                        # Employee names management
                        st.subheader("✏️ Employee Names Management")
                        st.write("**Edit employee names for each biometric ID:**")
                        
                        # Create columns for name inputs
                        cols = st.columns(2)
                        employee_names = {}
                        
                        for idx, emp_id in enumerate(employees_in_month):
                            with cols[idx % 2]:
                                # Get current name
                                current_settings = st.session_state.employee_settings.get(
                                    emp_id, 
                                    {"name": f"EMPLOYEE {emp_id}", "employee_no": emp_id}
                                )
                                current_name = current_settings["name"]
                                
                                # Create input field
                                new_name = st.text_input(
                                    f"**ID: {emp_id}**",
                                    value=current_name,
                                    key=f"name_{emp_id}_{month_name}_{year_num}",
                                    help=f"Enter name for employee with biometric ID {emp_id}"
                                )
                                
                                # Save immediately to session state
                                if new_name.strip():
                                    st.session_state.employee_settings[emp_id] = {
                                        "name": new_name.strip().upper(),
                                        "employee_no": emp_id
                                    }
                                    employee_names[emp_id] = new_name.strip().upper()
                        
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
                                        emp_df = month_df[month_df["EmployeeNo"] == emp_id].copy()
                                        
                                        if emp_df.empty:
                                            errors.append(f"❌ {emp_id}: No attendance data found")
                                            continue
                                        
                                        # Get employee name
                                        emp_settings = st.session_state.employee_settings.get(
                                            emp_id, 
                                            {"name": f"EMPLOYEE {emp_id}", "employee_no": emp_id}
                                        )
                                        emp_name = emp_settings["name"]
                                        
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
                                            "employee_no": emp_id,
                                            "employee_name": emp_name,
                                            "pdf_file": pdf_file
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
                                    st.markdown(f"""
                                    <div class="success-box">
                                        <strong>✅ Successfully generated {success_count} DTR PDF files!</strong>
                                    </div>
                                    """, unsafe_allow_html=True)
                                    
                                    # Show Richard P. Samoranos special note
                                    richard_pdf = next((f for f in pdf_files if f["employee_no"] == "7220970"), None)
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
                                                display_name = file_info["employee_name"]
                                                if len(display_name) > 25:
                                                    display_name = display_name[:22] + "..."
                                                
                                                # Special styling for Richard
                                                if file_info["employee_no"] == "7220970":
                                                    st.success(f"✅ **{display_name}** (ID: {file_info['employee_no']})")
                                                else:
                                                    st.write(f"{display_name} (ID: {file_info['employee_no']})")
                                                
                                                # Download button
                                                st.download_button(
                                                    label=f"⬇️ Download {display_name}",
                                                    data=file_info["pdf_file"],
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
    <div style='text-align: center; color: #6B7280; padding: 20px;'>
        <p><strong>Civil Service Form No. 48 - DTR Generator</strong><br>
        <small>Version 4.0 | Manual National High School | Division of Davao del Sur</small></p>
    </div>
""", unsafe_allow_html=True)

# Installation instructions in sidebar
with st.sidebar:
    st.title("📋 About This App")
    
    st.markdown("""
    ### Key Features:
    ✅ **Exact Template Match** - Follows Civil Service Form 48 format
    ✅ **PDF Output** - Generates professional PDF files
    ✅ **Auto Saturdays/Sundays** - Correctly marks weekends
    ✅ **Undertime Calculation** - Accurate hour/minute calculation
    ✅ **Richard P. Samoranos** - Pre-loaded (Employee No. 7220970)
    ✅ **Batch Processing** - Generate all employees at once
    
    ### Expected File Format:
    ```
    7220970 2025-12-01 07:30:00
    7220970 2025-12-01 11:50:00
    7220970 2025-12-01 12:50:00
    7220970 2025-12-01 16:30:00
    ```
    
    ### Contact:
    For issues or questions, please contact the IT Department.
    """)
