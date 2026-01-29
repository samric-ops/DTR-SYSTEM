import streamlit as st
import pandas as pd
import calendar
from reportlab.lib.pagesizes import A4, portrait, landscape
from reportlab.lib.units import cm, inch, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO
import math

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
if st.button("📄 Generate DTR PDF File", type="primary"):
    try:
        # Create buffer for PDF
        buffer = BytesIO()
        
        # Page setup - A4 Portrait with custom margins
        page_width, page_height = portrait(A4)
        
        # Margins in cm (convert to points: 1 cm = 28.3465 points)
        top_margin = 1.9 * cm
        bottom_margin = 0.49 * cm
        left_margin = 1.27 * cm
        right_margin = 1.27 * cm
        
        # Calculate usable width and height
        usable_width = page_width - left_margin - right_margin
        usable_height = page_height - top_margin - bottom_margin
        
        # Each DTR gets half of usable width
        dtr_width = usable_width / 2
        spacer_width = 0.5 * cm  # Space between two DTRs
        
        # Create PDF document
        doc = SimpleDocTemplate(
            buffer,
            pagesize=portrait(A4),
            leftMargin=left_margin,
            rightMargin=right_margin,
            topMargin=top_margin,
            bottomMargin=bottom_margin
        )
        
        # Styles
        styles = getSampleStyleSheet()
        
        # Custom styles
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=12,
            alignment=1,  # Center
            spaceAfter=6
        )
        
        header_style = ParagraphStyle(
            'CustomHeader',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=8,
            alignment=1,  # Center
            spaceAfter=3
        )
        
        name_style = ParagraphStyle(
            'CustomName',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=10,
            alignment=1,  # Center
            spaceAfter=3,
            textDecoration='underline'
        )
        
        small_style = ParagraphStyle(
            'CustomSmall',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=7,
            alignment=1,  # Center
            spaceAfter=2
        )
        
        table_header_style = ParagraphStyle(
            'TableHeader',
            parent=styles['Normal'],
            fontName='Helvetica-Bold',
            fontSize=8,
            alignment=1,  # Center
        )
        
        table_cell_style = ParagraphStyle(
            'TableCell',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=8,
            alignment=1,  # Center
        )
        
        footer_style = ParagraphStyle(
            'FooterStyle',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=8,
            alignment=1,  # Center
            fontStyle='italic'
        )
        
        # -------- FUNCTION TO CREATE ONE DTR TABLE --------
        def create_dtr_table_data(is_left=True):
            """Create table data for one DTR"""
            table_data = []
            
            # 1. REPUBLIC Header (4 lines)
            table_data.append([Paragraph(
                "REPUBLIC OF THE PHILIPPINES<br/>"
                "Department of Education<br/>"
                "Division of Davao del Sur<br/>"
                "Manual National High School",
                header_style
            )])
            
            # 2. Blank line
            table_data.append([""])
            
            # 3. Civil Service Form and Employee No.
            form_text = f"Civil Service Form No. 48{'&nbsp;' * 30}Employee No. {employee_no}"
            table_data.append([Paragraph(form_text, header_style)])
            
            # 4. Blank line
            table_data.append([""])
            
            # 5. DAILY TIME RECORD
            table_data.append([Paragraph("DAILY TIME RECORD", title_style)])
            
            # 6. -----o0o-----
            table_data.append([Paragraph("-----o0o-----", header_style)])
            
            # 7. Employee Name
            table_data.append([Paragraph(employee_name, name_style)])
            
            # 8. (Name) label
            table_data.append([Paragraph("(Name)", small_style)])
            
            # 9. Period and Official Hours
            period_text = f"For the month of {month.upper()} {year}"
            hours_text = f"Official hours for arrival and departure<br/>" \
                        f"Regular days: {am_hours} / {pm_hours}<br/>" \
                        f"Saturdays: {saturday_hours}"
            
            table_data.append([
                Paragraph(period_text, header_style),
                "",
                "",
                "",
                "",
                Paragraph(hours_text, header_style)
            ])
            
            # 10. Blank line before table
            table_data.append([""])
            
            # 11. TABLE HEADER - First row
            table_data.append([
                Paragraph("Day", table_header_style),
                Paragraph("A.M.", table_header_style),
                "",
                Paragraph("P.M.", table_header_style),
                "",
                Paragraph("Undertime", table_header_style),
                ""
            ])
            
            # 12. TABLE HEADER - Second row
            table_data.append([
                "",  # Day stays merged
                Paragraph("Arrival", table_header_style),
                Paragraph("Departure", table_header_style),
                Paragraph("Arrival", table_header_style),
                Paragraph("Departure", table_header_style),
                Paragraph("Hours", table_header_style),
                Paragraph("Minutes", table_header_style)
            ])
            
            # 13-43. TABLE DATA (31 days)
            for day_num in range(1, 32):
                row = []
                
                if day_num <= len(edited_df):
                    row_data = edited_df.iloc[day_num - 1]
                    
                    # Day number
                    row.append(Paragraph(str(day_num), table_cell_style))
                    
                    # Check if SATURDAY/SUNDAY
                    if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                        # Merge cells for weekend
                        row.append(Paragraph(str(row_data["AM In"]).strip(), table_cell_style))
                        row.append("")  # Will be merged
                        row.append("")  # Will be merged
                        row.append("")  # Will be merged
                    else:
                        # Regular day
                        row.append(Paragraph(
                            "" if pd.isna(row_data["AM In"]) else str(row_data["AM In"]), 
                            table_cell_style
                        ))
                        row.append(Paragraph(
                            "" if pd.isna(row_data["AM Out"]) else str(row_data["AM Out"]), 
                            table_cell_style
                        ))
                        row.append(Paragraph(
                            "" if pd.isna(row_data["PM In"]) else str(row_data["PM In"]), 
                            table_cell_style
                        ))
                        row.append(Paragraph(
                            "" if pd.isna(row_data["PM Out"]) else str(row_data["PM Out"]), 
                            table_cell_style
                        ))
                    
                    # Undertime columns (empty)
                    row.append("")  # Hours
                    row.append("")  # Minutes
                else:
                    # Empty row
                    row = ["", "", "", "", "", "", ""]
                
                table_data.append(row)
            
            # 44. TOTAL ROW
            table_data.append([
                Paragraph("TOTAL", table_header_style),
                "", "", "", "", "", ""
            ])
            
            # 45-46. Blank lines
            table_data.append([""])
            table_data.append([""])
            
            # 47. Certification
            cert_text = "I certify on my honor that the above is a true and correct report " \
                       "of the hours of work performed, record of which was made daily at " \
                       "the time of arrival and departure from office."
            table_data.append([Paragraph(cert_text, footer_style)])
            
            # 48. Blank line
            table_data.append([""])
            
            # 49. Signature line
            table_data.append([Paragraph("_________________________", header_style)])
            
            # 50. Blank line
            table_data.append([""])
            
            # 51. VERIFIED text
            table_data.append([Paragraph("VERIFIED as to the prescribed office hours:", header_style)])
            
            # 52. Blank line
            table_data.append([""])
            
            # 53. Principal signature line
            table_data.append([Paragraph("_________________________", header_style)])
            
            # 54. Principal III
            table_data.append([Paragraph("Principal III", table_header_style)])
            
            return table_data
        
        # -------- CREATE TWO DTR TABLES --------
        # We'll create a master table that contains both DTRs side by side
        
        # Create data for left DTR
        left_data = create_dtr_table_data(is_left=True)
        right_data = create_dtr_table_data(is_left=False)
        
        # Combine into one table with two columns
        # Each DTR has 7 columns, so total 14 columns + 1 spacer
        
        # Calculate column widths
        # Left DTR: 7 columns
        # Spacer: 1 column
        # Right DTR: 7 columns
        total_cols = 15  # 7 + 1 + 7
        
        # Column widths (in points)
        col_widths = []
        
        # Left DTR columns
        left_cols = [dtr_width * 0.12,   # Day
                    dtr_width * 0.16,    # A.M. Arrival
                    dtr_width * 0.16,    # A.M. Departure
                    dtr_width * 0.16,    # P.M. Arrival
                    dtr_width * 0.16,    # P.M. Departure
                    dtr_width * 0.12,    # Hours
                    dtr_width * 0.12]    # Minutes
        
        col_widths.extend(left_cols)
        col_widths.append(spacer_width)  # Spacer
        col_widths.extend(left_cols)     # Right DTR (same widths)
        
        # Combine data from both DTRs
        combined_data = []
        
        # Process each row
        for i in range(max(len(left_data), len(right_data))):
            combined_row = []
            
            # Add left DTR row
            if i < len(left_data):
                left_row = left_data[i]
                if isinstance(left_row, list):
                    combined_row.extend(left_row)
                else:
                    # Single cell that should span all 7 columns
                    combined_row.append(left_row)
                    combined_row.extend([""] * 6)  # Fill remaining cells
            else:
                combined_row.extend([""] * 7)
            
            # Add spacer column
            combined_row.append("")
            
            # Add right DTR row
            if i < len(right_data):
                right_row = right_data[i]
                if isinstance(right_row, list):
                    combined_row.extend(right_row)
                else:
                    # Single cell that should span all 7 columns
                    combined_row.append(right_row)
                    combined_row.extend([""] * 6)  # Fill remaining cells
            else:
                combined_row.extend([""] * 7)
            
            combined_data.append(combined_row)
        
        # Create the table
        table = Table(combined_data, colWidths=col_widths)
        
        # Apply table styles
        table.setStyle(TableStyle([
            # Global settings
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            
            # Header rows (0-8) - span across 7 columns for each DTR
            ('SPAN', (0, 0), (6, 0)),   # REPUBLIC header left
            ('SPAN', (8, 0), (14, 0)),  # REPUBLIC header right
            
            ('SPAN', (0, 2), (6, 2)),   # Civil Service Form left
            ('SPAN', (8, 2), (14, 2)),  # Civil Service Form right
            
            ('SPAN', (0, 4), (6, 4)),   # DAILY TIME RECORD left
            ('SPAN', (8, 4), (14, 4)),  # DAILY TIME RECORD right
            
            ('SPAN', (0, 5), (6, 5)),   # -----o0o----- left
            ('SPAN', (8, 5), (14, 5)),  # -----o0o----- right
            
            ('SPAN', (0, 6), (6, 6)),   # Employee Name left
            ('SPAN', (8, 6), (14, 6)),  # Employee Name right
            
            ('SPAN', (0, 7), (6, 7)),   # (Name) left
            ('SPAN', (8, 7), (14, 7)),  # (Name) right
            
            # Period and hours row (row 8)
            ('SPAN', (0, 8), (4, 8)),   # Period left
            ('SPAN', (5, 8), (6, 8)),   # Hours left
            ('SPAN', (8, 8), (12, 8)),  # Period right
            ('SPAN', (13, 8), (14, 8)), # Hours right
            
            # Table headers (rows 9-10)
            ('SPAN', (0, 9), (0, 10)),  # Day column left
            ('SPAN', (1, 9), (2, 9)),   # A.M. left
            ('SPAN', (3, 9), (4, 9)),   # P.M. left
            ('SPAN', (5, 9), (6, 9)),   # Undertime left
            
            ('SPAN', (8, 9), (8, 10)),  # Day column right
            ('SPAN', (9, 9), (10, 9)),  # A.M. right
            ('SPAN', (11, 9), (12, 9)), # P.M. right
            ('SPAN', (13, 9), (14, 9)), # Undertime right
            
            # Merge SATURDAY/SUNDAY cells
        ]))
        
        # Add specific spans for weekend days
        for day_num in range(1, 32):
            row_idx = 10 + day_num  # Starting from row 11 (0-indexed)
            
            if day_num <= len(edited_df):
                row_data = edited_df.iloc[day_num - 1]
                
                if str(row_data["AM In"]).strip() in ["SATURDAY", "SUNDAY"]:
                    # Left DTR
                    table.setStyle(TableStyle([
                        ('SPAN', (1, row_idx), (4, row_idx)),  # Merge A.M. and P.M. columns
                    ]))
                    # Right DTR
                    table.setStyle(TableStyle([
                        ('SPAN', (9, row_idx), (12, row_idx)),  # Merge A.M. and P.M. columns
                    ]))
        
        # TOTAL row spans
        total_row_idx = 10 + 31 + 1  # After 31 days
        table.setStyle(TableStyle([
            ('SPAN', (0, total_row_idx), (4, total_row_idx)),  # TOTAL left
            ('SPAN', (8, total_row_idx), (12, total_row_idx)), # TOTAL right
        ]))
        
        # Footer spans
        footer_start = total_row_idx + 3
        table.setStyle(TableStyle([
            # Certification
            ('SPAN', (0, footer_start), (6, footer_start)),    # Left
            ('SPAN', (8, footer_start), (14, footer_start)),   # Right
            
            # Signature line
            ('SPAN', (0, footer_start + 2), (6, footer_start + 2)),    # Left
            ('SPAN', (8, footer_start + 2), (14, footer_start + 2)),   # Right
            
            # VERIFIED text
            ('SPAN', (0, footer_start + 4), (6, footer_start + 4)),    # Left
            ('SPAN', (8, footer_start + 4), (14, footer_start + 4)),   # Right
            
            # Principal signature line
            ('SPAN', (0, footer_start + 6), (6, footer_start + 6)),    # Left
            ('SPAN', (8, footer_start + 6), (14, footer_start + 6)),   # Right
            
            # Principal III
            ('SPAN', (0, footer_start + 7), (6, footer_start + 7)),    # Left
            ('SPAN', (8, footer_start + 7), (14, footer_start + 7)),   # Right
        ]))
        
        # Add borders to table cells
        border_style = ('GRID', (0, 9), (6, total_row_idx), 0.5, colors.black)  # Left DTR table
        table.setStyle(TableStyle([border_style]))
        
        border_style_right = ('GRID', (8, 9), (14, total_row_idx), 0.5, colors.black)  # Right DTR table
        table.setStyle(TableStyle([border_style_right]))
        
        # Add border to spacer column to separate DTRs
        table.setStyle(TableStyle([
            ('LINEAFTER', (6, 0), (6, -1), 0.5, colors.white),  # White line as spacer
        ]))
        
        # Build PDF
        elements = []
        elements.append(table)
        
        # Build the PDF
        doc.build(elements)
        
        # Get PDF data
        buffer.seek(0)
        pdf_data = buffer.getvalue()
        
        st.success("✅ DTR PDF file generated successfully!")
        
        # Create safe filename
        safe_name = "".join([c if c.isalnum() or c in "._- " else "_" for c in employee_name])
        
        st.download_button(
            "📥 Download PDF File",
            pdf_data,
            file_name=f"DTR_CSForm48_{safe_name}_{month}_{year}.pdf",
            mime="application/pdf"
        )
        
        st.info("""
        **📝 Document Specifications:**
        - Format: **PDF (Portable Document Format)**
        - Paper: **A4 Portrait**
        - Margins: Top 1.9cm, Bottom 0.49cm, Left/Right 1.27cm
        - Two identical DTRs side by side
        - Each DTR contains complete 1-31 days
        - Proper table borders and formatting
        - Ready to print
        """)
        
    except Exception as e:
        st.error(f"❌ Error generating PDF file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
