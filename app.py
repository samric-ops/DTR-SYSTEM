import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io
import base64

# Page configuration
st.set_page_config(
    page_title="DTR System - Attendance Log Reader",
    page_icon="📊",
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
    .sub-header {
        font-size: 1.5rem;
        color: #3B82F6;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .stButton button {
        background-color: #3B82F6;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 1rem;
    }
    .success-box {
        background-color: #D1FAE5;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #10B981;
    }
    .info-box {
        background-color: #DBEAFE;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #3B82F6;
    }
</style>
""", unsafe_allow_html=True)

# App title
st.markdown('<h1 class="main-header">📊 DTR Attendance Log System</h1>', unsafe_allow_html=True)
st.markdown("Upload your `.dat` attendance files and analyze employee time records.")

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None

# Sidebar
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3448/3448512.png", width=100)
    st.title("Navigation")
    
    st.markdown("---")
    
    st.subheader("📤 Upload File")
    uploaded_file = st.file_uploader(
        "Choose a DAT file",
        type=['dat', 'txt', 'csv'],
        help="Upload your attendance log file"
    )
    
    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file
        
        # Read the file
        try:
            # Try to detect delimiter
            content = uploaded_file.getvalue().decode('utf-8')
            first_line = content.split('\n')[0].strip()
            
            if '\t' in first_line:
                delimiter = '\t'
            elif ',' in first_line:
                delimiter = ','
            else:
                delimiter = ' '
            
            # Reset file pointer
            uploaded_file.seek(0)
            
            # Read based on file size
            if uploaded_file.size > 10 * 1024 * 1024:  # 10MB
                st.info("Large file detected. Reading in chunks...")
                chunks = []
                for chunk in pd.read_csv(uploaded_file, header=None, delimiter=delimiter, chunksize=10000):
                    chunks.append(chunk)
                df = pd.concat(chunks, ignore_index=True)
            else:
                df = pd.read_csv(uploaded_file, header=None, delimiter=delimiter)
            
            # Assign column names
            column_names = ['UserID', 'DateTime', 'Status', 'Verification']
            for i in range(min(len(df.columns), len(column_names))):
                df = df.rename(columns={i: column_names[i]})
            
            # Convert DateTime
            df['DateTime'] = pd.to_datetime(df['DateTime'], errors='coerce')
            
            # Add derived columns
            df['Date'] = df['DateTime'].dt.date
            df['Time'] = df['DateTime'].dt.time
            df['Day'] = df['DateTime'].dt.day_name()
            df['Hour'] = df['DateTime'].dt.hour
            df['Month'] = df['DateTime'].dt.month_name()
            
            # Sort by DateTime
            df = df.sort_values('DateTime')
            
            # Reset index
            df.reset_index(drop=True, inplace=True)
            
            st.session_state.df = df
            
            st.success(f"✅ File loaded successfully!")
            st.info(f"**Records:** {len(df):,} | **Users:** {df['UserID'].nunique()} | **Date Range:** {df['Date'].min()} to {df['Date'].max()}")
            
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
    
    st.markdown("---")
    
    if st.session_state.df is not None:
        st.subheader("⚙️ Settings")
        date_range = st.date_input(
            "Select Date Range",
            value=[st.session_state.df['Date'].min(), st.session_state.df['Date'].max()],
            min_value=st.session_state.df['Date'].min(),
            max_value=st.session_state.df['Date'].max()
        )
        
        user_ids = st.multiselect(
            "Select Users",
            options=sorted(st.session_state.df['UserID'].unique()),
            default=sorted(st.session_state.df['UserID'].unique())[:5] if len(st.session_state.df['UserID'].unique()) > 0 else []
        )
        
        st.markdown("---")
    
    st.subheader("📞 Help")
    st.markdown("""
    **File Format:**
    - UserID,DateTime,Status,Verification
    - Example: `1,2024-01-01 08:00:00,0,1`
    
    **Support:** contact@example.com
    """)

# Main content
if st.session_state.df is not None:
    df = st.session_state.df
    
    # Apply filters
    if 'date_range' in locals() and date_range:
        if len(date_range) == 2:
            mask = (df['Date'] >= date_range[0]) & (df['Date'] <= date_range[1])
            df = df[mask]
    
    if 'user_ids' in locals() and user_ids:
        df = df[df['UserID'].isin(user_ids)]
    
    # Dashboard metrics
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Records", f"{len(df):,}")
    
    with col2:
        st.metric("Unique Users", df['UserID'].nunique())
    
    with col3:
        st.metric("Date Range", f"{df['Date'].min()} to {df['Date'].max()}")
    
    with col4:
        avg_records = len(df) / df['UserID'].nunique() if df['UserID'].nunique() > 0 else 0
        st.metric("Avg Records/User", f"{avg_records:.1f}")
    
    # Tabs for different views
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📋 Data View", "📈 Analytics", "⏱️ Hours Calculation", "📊 Reports", "💾 Export"])
    
    with tab1:
        st.markdown('<h3 class="sub-header">Raw Data</h3>', unsafe_allow_html=True)
        
        # Search and filter
        col1, col2 = st.columns(2)
        with col1:
            search_user = st.text_input("Search User ID", "")
        with col2:
            items_per_page = st.selectbox("Rows per page", [10, 20, 50, 100], index=0)
        
        # Apply search filter
        if search_user:
            df_display = df[df['UserID'].astype(str).str.contains(search_user, case=False)]
        else:
            df_display = df.copy()
        
        # Pagination
        total_pages = max(1, len(df_display) // items_per_page)
        page_number = st.number_input("Page", min_value=1, max_value=total_pages, value=1)
        
        start_idx = (page_number - 1) * items_per_page
        end_idx = start_idx + items_per_page
        
        # Display table
        st.dataframe(
            df_display[['UserID', 'Date', 'Time', 'Day', 'Status', 'Verification']]
            .iloc[start_idx:end_idx],
            use_container_width=True,
            height=400
        )
        
        st.caption(f"Showing {start_idx+1}-{min(end_idx, len(df_display))} of {len(df_display)} records")
    
    with tab2:
        st.markdown('<h3 class="sub-header">Analytics Dashboard</h3>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Records per user
            user_counts = df['UserID'].value_counts().reset_index()
            user_counts.columns = ['UserID', 'Count']
            
            fig1 = px.bar(
                user_counts.head(20),
                x='UserID',
                y='Count',
                title="Top 20 Users by Record Count",
                color='Count',
                color_continuous_scale='Blues'
            )
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            # Records by day of week
            day_counts = df['Day'].value_counts().reset_index()
            day_counts.columns = ['Day', 'Count']
            
            # Order days properly
            day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            day_counts['Day'] = pd.Categorical(day_counts['Day'], categories=day_order, ordered=True)
            day_counts = day_counts.sort_values('Day')
            
            fig2 = px.bar(
                day_counts,
                x='Day',
                y='Count',
                title="Records by Day of Week",
                color='Count',
                color_continuous_scale='Greens'
            )
            st.plotly_chart(fig2, use_container_width=True)
        
        # Hourly distribution
        hourly_data = df.groupby('Hour').size().reset_index(name='Count')
        fig3 = px.line(
            hourly_data,
            x='Hour',
            y='Count',
            title="Hourly Distribution of Logs",
            markers=True
        )
        fig3.update_layout(xaxis=dict(tickmode='linear', dtick=1))
        st.plotly_chart(fig3, use_container_width=True)
    
    with tab3:
        st.markdown('<h3 class="sub-header">Working Hours Calculation</h3>', unsafe_allow_html=True)
        
        st.info("""
        This calculation pairs IN and OUT records for each user per day.
        Assumption: First record of the day = IN, Last record = OUT
        """)
        
        # Calculate working hours
        try:
            # Group by user and date
            working_hours = []
            
            for (user_id, date), group in df.groupby(['UserID', 'Date']):
                if len(group) >= 2:
                    first_log = group['DateTime'].min()
                    last_log = group['DateTime'].max()
                    
                    # Calculate hours worked
                    hours_worked = (last_log - first_log).total_seconds() / 3600
                    
                    working_hours.append({
                        'UserID': user_id,
                        'Date': date,
                        'First Log': first_log.time(),
                        'Last Log': last_log.time(),
                        'Records': len(group),
                        'Hours Worked': round(hours_worked, 2)
                    })
            
            if working_hours:
                hours_df = pd.DataFrame(working_hours)
                
                # Summary by user
                user_summary = hours_df.groupby('UserID').agg({
                    'Hours Worked': 'sum',
                    'Records': 'sum',
                    'Date': 'count'
                }).round(2)
                user_summary.columns = ['Total Hours', 'Total Records', 'Days Worked']
                user_summary['Avg Hours/Day'] = (user_summary['Total Hours'] / user_summary['Days Worked']).round(2)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Daily Working Hours")
                    st.dataframe(hours_df.sort_values(['UserID', 'Date']), use_container_width=True)
                
                with col2:
                    st.subheader("User Summary")
                    st.dataframe(user_summary, use_container_width=True)
                    
                    # Visualization
                    fig = px.bar(
                        user_summary.reset_index(),
                        x='UserID',
                        y='Total Hours',
                        title="Total Working Hours by User",
                        color='Total Hours',
                        color_continuous_scale='Viridis'
                    )
                    st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("Not enough data to calculate working hours. Need at least 2 records per user per day.")
        
        except Exception as e:
            st.error(f"Error calculating hours: {str(e)}")
    
    with tab4:
        st.markdown('<h3 class="sub-header">Reports</h3>', unsafe_allow_html=True)
        
        report_type = st.selectbox(
            "Select Report Type",
            ["User Activity Report", "Attendance Summary", "Daily Log Report", "Custom Report"]
        )
        
        if report_type == "User Activity Report":
            selected_user = st.selectbox("Select User", df['UserID'].unique())
            user_data = df[df['UserID'] == selected_user]
            
            if not user_data.empty:
                st.subheader(f"Activity Report for User {selected_user}")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Total Records", len(user_data))
                    st.metric("First Record", user_data['Date'].min())
                    st.metric("Last Record", user_data['Date'].max())
                
                with col2:
                    days_active = user_data['Date'].nunique()
                    st.metric("Days Active", days_active)
                    st.metric("Most Active Day", user_data['Day'].mode().iloc[0] if not user_data['Day'].mode().empty else "N/A")
                
                # User's daily pattern
                user_daily = user_data.groupby('Hour').size().reset_index(name='Count')
                fig = px.line(user_daily, x='Hour', y='Count', title=f"User {selected_user} - Hourly Pattern")
                st.plotly_chart(fig, use_container_width=True)
        
        elif report_type == "Attendance Summary":
            # Generate summary report
            summary_data = []
            
            for user_id in df['UserID'].unique():
                user_df = df[df['UserID'] == user_id]
                summary_data.append({
                    'UserID': user_id,
                    'Total Records': len(user_df),
                    'Days Active': user_df['Date'].nunique(),
                    'First Date': user_df['Date'].min(),
                    'Last Date': user_df['Date'].max(),
                    'Most Common Hour': user_df['Hour'].mode().iloc[0] if not user_df['Hour'].mode().empty else "N/A"
                })
            
            summary_df = pd.DataFrame(summary_data)
            st.dataframe(summary_df, use_container_width=True)
    
    with tab5:
        st.markdown('<h3 class="sub-header">Export Data</h3>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📥 Download as CSV", use_container_width=True):
                csv = df.to_csv(index=False)
                b64 = base64.b64encode(csv.encode()).decode()
                href = f'<a href="data:file/csv;base64,{b64}" download="attendance_data.csv">Click here to download CSV</a>'
                st.markdown(href, unsafe_allow_html=True)
        
        with col2:
            if st.button("📥 Download as Excel", use_container_width=True):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Attendance')
                
                b64 = base64.b64encode(output.getvalue()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="attendance_data.xlsx">Click here to download Excel</a>'
                st.markdown(href, unsafe_allow_html=True)
        
        with col3:
            if st.button("📄 Generate Summary Report", use_container_width=True):
                # Create a text report
                report = f"""
                DTR ATTENDANCE REPORT
                ======================
                Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                
                SUMMARY STATISTICS:
                - Total Records: {len(df):,}
                - Unique Users: {df['UserID'].nunique()}
                - Date Range: {df['Date'].min()} to {df['Date'].max()}
                - Most Active Day: {df['Day'].mode().iloc[0] if not df['Day'].mode().empty else 'N/A'}
                
                USER SUMMARY:
                """
                
                for user_id in sorted(df['UserID'].unique())[:10]:  # Top 10 users
                    user_data = df[df['UserID'] == user_id]
                    report += f"\n- User {user_id}: {len(user_data)} records, {user_data['Date'].nunique()} days active"
                
                st.download_button(
                    label="📥 Download Report",
                    data=report,
                    file_name="attendance_summary.txt",
                    mime="text/plain",
                    use_container_width=True
                )
        
        st.markdown("---")
        
        # Preview export data
        st.subheader("Preview Export Data")
        st.dataframe(df.head(100), use_container_width=True)

else:
    # Welcome screen when no data loaded
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.image("https://cdn-icons-png.flaticon.com/512/3448/3448579.png", width=200)
        st.markdown("""
        <div class="info-box">
        <h3>📤 Upload your DAT file to get started</h3>
        <p>Supported formats:</p>
        <ul>
            <li>ZKTeco/ZKTime DAT files</li>
            <li>CSV files with attendance data</li>
            <li>Text files with timestamp data</li>
        </ul>
        <p><strong>Expected format:</strong> UserID,DateTime,Status,Verification</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Sample data download
        st.markdown("---")
        st.subheader("Need a sample file?")
        
        sample_data = """1,2024-01-01 08:00:00,0,1
1,2024-01-01 17:00:00,0,1
2,2024-01-01 08:15:00,0,1
2,2024-01-01 17:30:00,0,1
1,2024-01-02 08:05:00,0,1
1,2024-01-02 16:55:00,0,1"""
        
        st.download_button(
            label="📥 Download Sample DAT File",
            data=sample_data,
            file_name="sample_attendance.dat",
            mime="text/plain",
            use_container_width=True
        )

# Footer
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: gray;">
    <p>DTR Attendance System v1.0 | Built with Streamlit | For educational purposes</p>
    </div>
    """,
    unsafe_allow_html=True
)
