import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

# Page Configuration
st.set_page_config(page_title="DTR Excel Generator", layout="wide")

st.title("📋 DTR Excel Generator")
st.markdown("---")

# Input Form sa Sidebar
with st.sidebar:
    st.header("Employee Information")
    
    employee_name = st.text_input("Employee Name", "SAMORANOS, RICHARD P.")
    month_year = st.text_input("Month and Year", "DECEMBER 2025")
    
    st.header("Office Hours")
    am_in = st.text_input("AM In", "07:30")
    am_out = st.text_input("AM Out", "11:50")
    pm_in = st.text_input("PM In", "12:50")
    pm_out = st.text_input("PM Out", "16:30")
    saturday_hours = st.text_input("Saturday Hours", "AS REQUIRED")
    
    st.header("Generate DTR")
    if st.button("Generate Excel File", type="primary"):
        st.session_state.generate = True

# Main Content
if 'generate' not in st.session_state:
    st.session_state.generate = False

if st.session_state.generate:
    # Create Excel file in memory
    output = BytesIO()
    
    # Create a Pandas Excel writer using BytesIO
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Create DataFrame para sa header information
        header_data = {
            'REPUBLIC OF THE PHILIPPINES': [''],
            'Department of Education': [''],
            'Division of Davao del Sur': [''],
            'MANUAL NATIONAL HIGH SCHOOL': [''],
            '': [''],
            'DAILY TIME RECORD': [''],
            '---000---': [''],
            '': [''],
            'Employee Name': [employee_name],
            'For the month of': [month_year],
            '': [''],
            'Official hours for arrival and departure': [''],
            'Regular days': [f'{am_in} - {am_out} / {pm_in} - {pm_out}'],
            'Saturday': [saturday_hours],
            '': [''],
        }
        
        header_df = pd.DataFrame(header_data)
        header_df.to_excel(writer, sheet_name='DTR', index=False, startrow=0, startcol=0)
        
        # Create the main table
        table_data = []
        
        # Column headers
        table_headers = [
            ['Day', 'A.M.', 'A.M.', 'P.M.', 'P.M.', 'Undertime', 'Undertime'],
            ['', 'Arrival', 'Departure', 'Arrival', 'Departure', 'Hours', 'Minutes']
        ]
        
        for headers in table_headers:
            table_data.append(headers)
        
        # Add data for each day of the month (assuming 31 days)
        for day in range(1, 32):
            row = [day, '', '', '', '', '', '']  # Empty row by default
            
            # Sample data - dito mo ilalagay ang actual time entries
            if day == 1:
                row = [1, '07:00', '11:46', '16:36', '16:36', '3', '14']
            elif day == 2:
                row = [2, '07:25', '11:48', '16:40', '16:40', '3', '37']
            elif day == 3:
                row = [3, '07:29', '11:46', '16:41', '16:41', '3', '43']
            elif day == 4:
                row = [4, '07:28', '11:49', '16:46', '16:46', '3', '39']
            elif day == 5:
                row = [5, '07:29', '11:41', '16:37', '16:37', '3', '48']
            elif day == 6:
                row = [6, 'SATURDAY', '', '', '', '', '']
            elif day == 7:
                row = [7, 'SUNDAY', '', '', '', '', '']
            elif day == 13:
                row = [13, 'SATURDAY', '', '', '', '', '']
            elif day == 14:
                row = [14, 'SUNDAY', '', '', '', '', '']
            elif day == 17:
                row = [17, '07:33', '11:41', '16:54', '16:54', '3', '52']
            
            table_data.append(row)
        
        # Add total row
        table_data.append(['TOTAL', '', '', '', '', '', ''])
        
        # Convert to DataFrame and write to Excel
        table_df = pd.DataFrame(table_data)
        table_df.to_excel(writer, sheet_name='DTR', index=False, header=False, startrow=len(header_df) + 1)
        
        # Add footer
        footer_data = {
            'Certification': [''],
            'I certify on my honor that the above is a true and correct report': [''],
            'of the hours of work performed, record of which was made daily': [''],
            'at the time of arrival and departure from office.': [''],
            '': [''],
            '___________________________________': [''],
            '(Signature of Employee)': [''],
            '': [''],
            'VERIFIED as to the prescribed office hours:': [''],
            '': [''],
            '___________________________________': [''],
            'Principal III': ['']
        }
        
        footer_df = pd.DataFrame(footer_data)
        footer_df.to_excel(writer, sheet_name='DTR', index=False, header=False, 
                          startrow=len(header_df) + len(table_df) + 3)
        
        # Adjust column widths
        worksheet = writer.sheets['DTR']
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column].width = adjusted_width
    
    # Get the Excel file data
    excel_data = output.getvalue()
    
    # Download button
    st.success("✅ DTR Excel file generated successfully!")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.download_button(
            label="📥 Download Excel File",
            data=excel_data,
            file_name=f"DTR_{employee_name.split(',')[0]}_{month_year.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    # Preview the data
    st.subheader("Preview of Generated DTR")
    
    # Display preview
    st.write("**Header Information:**")
    st.write(f"Employee: {employee_name}")
    st.write(f"Month: {month_year}")
    st.write(f"Regular Hours: {am_in} - {am_out} / {pm_in} - {pm_out}")
    st.write(f"Saturday: {saturday_hours}")
    
    st.write("\n**Time Entries (Sample Data):**")
    preview_df = pd.DataFrame(table_data[2:-1], columns=table_data[0])
    st.dataframe(preview_df, use_container_width=True)
    
    st.write("\n**Instructions:**")
    st.info("""
    1. Click the download button above to get the Excel file
    2. Open the file in Microsoft Excel
    3. Adjust formatting as needed (merge cells, borders, etc.)
    4. Print from Excel (Ctrl+P)
    """)
    
    # Reset button
    if st.button("Generate Another DTR"):
        st.session_state.generate = False
        st.rerun()

else:
    # Instructions page
    st.subheader("How to Use This DTR Generator")
    
    st.markdown("""
    ### 📝 Step-by-Step Guide:
    
    1. **Fill in the information** in the sidebar on the left
    2. **Click 'Generate Excel File'** button
    3. **Download** the generated Excel file
    4. **Open** in Microsoft Excel
    5. **Format** if needed (adjust column widths, add borders)
    6. **Print** from Excel
    
    ### 📋 Information Needed:
    - Employee Name
    - Month and Year
    - Regular Office Hours (AM In/Out, PM In/Out)
    - Saturday Hours
    
    ### ⚡ Quick Start:
    Just click the "Generate Excel File" button in the sidebar to use the sample data.
    """)
    
    # Sample Preview
    st.subheader("Sample Output Preview")
    
    sample_data = {
        'Day': ['1', '2', '3', '...', '17'],
        'AM Arrival': ['07:00', '07:25', '07:29', '...', '07:33'],
        'AM Departure': ['11:46', '11:48', '11:46', '...', '11:41'],
        'PM Arrival': ['16:36', '16:40', '16:41', '...', '16:54'],
        'PM Departure': ['16:36', '16:40', '16:41', '...', '16:54'],
        'Undertime Hrs': ['3', '3', '3', '...', '3'],
        'Undertime Min': ['14', '37', '43', '...', '52']
    }
    
    st.table(pd.DataFrame(sample_data))

# Footer
st.markdown("---")
st.caption("DTR Excel Generator v1.0 | Designed for Manual National High School")
