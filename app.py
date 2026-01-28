import streamlit as st
from datetime import datetime

# 1. Page Configuration (Para maging wide mode agad)
st.set_page_config(layout="wide", page_title="DTR Generator")

# 2. CSS para itago ang Streamlit elements kapag nag-print (Crucial ito!)
hide_streamlit_style = """
    <style>
        /* Itago ang menu, header, at footer ng Streamlit */
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        
        /* Ayusin ang padding para sumakop sa buong page */
        .block-container {
            padding-top: 0rem;
            padding-bottom: 0rem;
            padding-left: 0rem;
            padding-right: 0rem;
            max-width: 100%;
        }
        
        /* Siguraduhing A4 Landscape ang print layout */
        @media print {
            @page {
                size: A4 landscape;
                margin: 5mm;
            }
            body {
                transform: scale(1.0);
                margin: 0;
                padding: 0;
            }
            .dtr-box {
                break-inside: avoid;
            }
        }
        
        /* Force no scaling sa printing */
        body.print {
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# 3. Data mula sa image mo
# Ito ay sample data base sa iyong image
sample_data = {
    "employee_name": "SAMORANOS, RICHARD P.",
    "month_year": "DECEMBER 2025",
    "regular_hours": "07:30 – 11:50 / 12:50 – 16:30",
    "saturday_hours": "AS REQUIRED",
    
    # Time entries para sa bawat araw (base sa image mo)
    "time_entries": {
        1: {"am_arrival": "07:00", "am_departure": "11:46", "pm_arrival": "16:36", "pm_departure": "16:36", "undertime_hrs": "3", "undertime_min": "14"},
        2: {"am_arrival": "07:25", "am_departure": "11:48", "pm_arrival": "16:40", "pm_departure": "16:40", "undertime_hrs": "3", "undertime_min": "37"},
        3: {"am_arrival": "07:29", "am_departure": "11:46", "pm_arrival": "16:41", "pm_departure": "16:41", "undertime_hrs": "3", "undertime_min": "43"},
        4: {"am_arrival": "07:28", "am_departure": "11:49", "pm_arrival": "16:46", "pm_departure": "16:46", "undertime_hrs": "3", "undertime_min": "39"},
        5: {"am_arrival": "07:29", "am_departure": "11:41", "pm_arrival": "16:37", "pm_departure": "16:37", "undertime_hrs": "3", "undertime_min": "48"},
        6: {"am_arrival": "SATURDAY", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        7: {"am_arrival": "SUNDAY", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        8: {"am_arrival": "", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        9: {"am_arrival": "", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        10: {"am_arrival": "", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        11: {"am_arrival": "", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        12: {"am_arrival": "", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        13: {"am_arrival": "SATURDAY", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        14: {"am_arrival": "SUNDAY", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        15: {"am_arrival": "", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        16: {"am_arrival": "", "am_departure": "", "pm_arrival": "", "pm_departure": "", "undertime_hrs": "", "undertime_min": ""},
        17: {"am_arrival": "07:33", "am_departure": "11:41", "pm_arrival": "16:54", "pm_departure": "16:54", "undertime_hrs": "3", "undertime_min": "52"},
    }
}

# 4. Ang iyong HTML DTR Code
dtr_html_template_start = """
<style>
    body {
        font-family: "Arial", sans-serif;
        font-size: 10px;
        margin: 0;
        padding: 0;
        display: flex;
        justify-content: center;
        width: 100%;
        background-color: white;
    }
    .container {
        display: flex;
        width: 100%;
        gap: 20px;
        padding: 10px;
    }
    .dtr-box {
        flex: 1;
        padding: 5px 10px;
        border: 1px solid #ccc;
        background-color: white;
    }
    .header { text-align: center; margin-bottom: 5px; }
    .header h4, .header h3, .header p { margin: 2px 0; font-weight: normal; }
    .cs-form { font-style: italic; font-size: 9px; text-align: left; }
    .school-name { font-weight: bold; }
    .dtr-title { 
        font-weight: bold; 
        font-size: 14px; 
        margin: 10px 0 !important; 
        text-decoration: underline;
    }
    .employee-name {
        font-weight: bold; 
        text-transform: uppercase; 
        font-size: 14px;
        border-bottom: 1px solid black; 
        display: inline-block; 
        width: 100%; 
        margin-bottom: 10px;
        padding-bottom: 3px;
    }
    .details-row { 
        display: flex; 
        justify-content: space-between; 
        margin-bottom: 2px; 
        font-size: 9px;
    }
    .input-line { 
        border-bottom: 1px solid black; 
        flex-grow: 1; 
        margin-left: 5px; 
        text-align: center; 
        font-weight: bold;
        min-width: 200px;
    }
    
    /* Table Styling */
    table { 
        width: 100%; 
        border-collapse: collapse; 
        margin-top: 5px; 
        font-size: 9px; 
        table-layout: fixed;
    }
    th, td { 
        border: 1px solid black; 
        text-align: center; 
        padding: 2px; 
        height: 15px;
        word-wrap: break-word;
    }
    th {
        background-color: #f0f0f0;
        font-weight: bold;
    }
    
    /* Footer Styling */
    .footer { 
        margin-top: 15px; 
        text-align: left; 
        font-size: 9px; 
    }
    .certify { 
        margin-bottom: 20px; 
        font-style: italic;
        line-height: 1.2;
    }
    .signature-line { 
        border-top: 1px solid black; 
        width: 80%; 
        margin: 25px auto 5px auto; 
        text-align: center;
        font-size: 9px;
    }
    .signature-title {
        text-align: center;
        margin-top: 3px;
        font-size: 9px;
    }
    
    /* Para sa weekends */
    .weekend {
        font-weight: bold;
    }
    
    /* Para sa mga walang laman na araw */
    .empty-day {
        color: #666;
    }
</style>

<div class="container">
"""

# 5. Function para bumuo ng isang DTR copy
def build_dtr_copy(data, copy_number):
    html = f"""
    <div class="dtr-box">
        <div class="cs-form">Civil Service Form No. 48</div>
        <div class="header">
            <p>REPUBLIC OF THE PHILIPPINES</p>
            <p>Department of Education</p>
            <p>Division of Davao del Sur</p>
            <p class="school-name">MANUAL NATIONAL HIGH SCHOOL</p>
            <h3 class="dtr-title">DAILY TIME RECORD</h3>
            <p style="font-size: 8px; margin-top: -8px;">---000---</p>
        </div>
        <div class="header">
            <span class="employee-name">{data['employee_name']}</span>
            <p>(Name)</p>
        </div>
        <div class="details-row">
            <span>For the month of:</span>
            <span class="input-line">{data['month_year']}</span>
        </div>
        <div class="details-row"><span>Official hours for arrival and departure:</span></div>
        <div class="details-row">
            <span>Regular days:</span>
            <span class="input-line">{data['regular_hours']}</span>
        </div>
        <div class="details-row">
            <span>Saturday:</span>
            <span class="input-line">{data['saturday_hours']}</span>
        </div>

        <table>
            <thead>
                <tr>
                    <th rowspan="2" style="width: 5%;">Day</th>
                    <th colspan="2" style="width: 25%;">A.M.</th>
                    <th colspan="2" style="width: 25%;">P.M.</th>
                    <th colspan="2" style="width: 15%;">Undertime</th>
                </tr>
                <tr>
                    <th style="width: 12.5%;">Arrival</th>
                    <th style="width: 12.5%;">Departure</th>
                    <th style="width: 12.5%;">Arrival</th>
                    <th style="width: 12.5%;">Departure</th>
                    <th style="width: 7.5%;">Hours</th>
                    <th style="width: 7.5%;">Minutes</th>
                </tr>
            </thead>
            <tbody>
    """
    
    # Generate rows para sa buong buwan (31 days)
    for day in range(1, 32):
        if day in data['time_entries']:
            entry = data['time_entries'][day]
            # Check kung weekend
            if entry['am_arrival'] in ['SATURDAY', 'SUNDAY']:
                html += f"""
                <tr>
                    <td>{day}</td>
                    <td colspan="6" class="weekend">{entry['am_arrival']}</td>
                </tr>
                """
            else:
                html += f"""
                <tr>
                    <td>{day}</td>
                    <td>{entry['am_arrival']}</td>
                    <td>{entry['am_departure']}</td>
                    <td>{entry['pm_arrival']}</td>
                    <td>{entry['pm_departure']}</td>
                    <td>{entry['undertime_hrs']}</td>
                    <td>{entry['undertime_min']}</td>
                </tr>
                """
        else:
            # Para sa mga walang data na araw
            html += f"""
            <tr>
                <td>{day}</td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
            </tr>
            """
    
    # Total row
    html += """
                <tr>
                    <td><strong>TOTAL</strong></td>
                    <td colspan="6"></td>
                </tr>
            </tbody>
        </table>

        <div class="footer">
            <p class="certify">I certify on my honor that the above is a true and correct report of the hours of work performed, record of which was made daily at the time of arrival and departure from office.</p>
            <div class="signature-line"></div>
            <div class="signature-title">(Signature of Employee)</div>
            
            <div style="margin-top: 25px;">
                <p>VERIFIED as to the prescribed office hours:</p>
                <br>
                <div style="text-align: center; border-top: 1px solid black; width: 60%; margin: 10px auto 0 auto;"></div>
                <div class="signature-title"><strong>Principal III</strong></div>
            </div>
        </div>
    </div>
    """
    
    return html

# 6. Build the complete HTML
full_html = dtr_html_template_start

# Unang kopya (Left side)
full_html += build_dtr_copy(sample_data, 1)

# Pangalawang kopya (Right side)
full_html += build_dtr_copy(sample_data, 2)

# Close container
full_html += "</div>"

# 7. Sidebar para sa user input (optional - pwedeng lagyan ng functionality)
with st.sidebar:
    st.header("DTR Configuration")
    
    # Pwede lagyan ng input fields dito para ma-edit ang data
    st.subheader("Employee Information")
    employee_name = st.text_input("Employee Name", value=sample_data["employee_name"])
    month_year = st.text_input("Month and Year", value=sample_data["month_year"])
    
    st.subheader("Office Hours")
    regular_hours = st.text_input("Regular Hours", value=sample_data["regular_hours"])
    saturday_hours = st.text_input("Saturday Hours", value=sample_data["saturday_hours"])
    
    if st.button("Update DTR"):
        # Update the sample data with user input
        sample_data["employee_name"] = employee_name
        sample_data["month_year"] = month_year
        sample_data["regular_hours"] = regular_hours
        sample_data["saturday_hours"] = saturday_hours
        st.rerun()
    
    st.divider()
    st.info("**Para i-print:**")
    st.write("1. Pindutin ang **Ctrl+P**")
    st.write("2. Piliin ang **Landscape** orientation")
    st.write("3. Piliin ang **A4** paper size")
    st.write("4. I-set ang margins sa **5mm**")
    st.write("5. I-uncheck ang **'Headers and footers'** option")

# 8. Render ang HTML
st.markdown(full_html, unsafe_allow_html=True)

# 9. Instructions sa baba
st.markdown("---")
st.info("**Note:** Ang DTR na ito ay base sa iyong provided image. Ang data para sa December 1-5 at 17 ay nakapaloob na.")
