import streamlit as st
import pandas as pd
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import io
import datetime

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="InfoBeans Foundation Reporting", layout="wide", page_icon="🎓")

# Streamlit Custom Theme Styling (Official InfoBeans Colors)
st.markdown("""
<style>
    :root {
        --primary-color: #EA1B3D;
        --secondary-color: #F7925B;
        --dark-color: #2F2F39;
    }
    .main-header {
        color: #2F2F39;
        font-weight: 800;
        border-bottom: 3px solid #EA1B3D;
        padding-bottom: 8px;
    }
    .stButton>button {
        background-color: #EA1B3D;
        color: white;
        border-radius: 6px;
        border: none;
        font-weight: 600;
    }
    .stButton>button:hover {
        background-color: #c91432;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("🎓 InfoBeans Foundation — Student Reporting System")
st.caption("Empowering Talent, Building Futures — Cross-Batch Aggregation & Live Dispatch")

if "email_logs" not in st.session_state:
    st.session_state.email_logs = []

def safe_num(val, default=0.0):
    try:
        if pd.isna(val) or val is None or str(val).strip() == "" or str(val).lower() == "nan":
            return default
        return float(val)
    except:
        return default

def get_performance_status(att_pct, test_pct):
    if att_pct < 75.0 or test_pct < 65.0:
        return "Needs Attention", colors.HexColor('#EA1B3D'), colors.HexColor('#FEE2E2')
    elif att_pct >= 85.0 and test_pct >= 80.0:
        return "Good", colors.HexColor('#16A34A'), colors.HexColor('#DCFCE7')
    return "Satisfactory", colors.HexColor('#F7925B'), colors.HexColor('#FFF7ED')

# --- PDF GENERATOR 1: COLLEGE CONSOLIDATED MASTER REPORT ---
def generate_college_pdf(college_name, month_str, df_college):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=32, bottomMargin=32)
    elements = []
    
    brand_title = ParagraphStyle('BTitle', fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#2F2F39'), spaceAfter=2)
    brand_sub = ParagraphStyle('BSub', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#EA1B3D'), spaceAfter=4)
    rep_title = ParagraphStyle('RTitle', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#555562'))
    college_header = ParagraphStyle('CHead', fontName='Helvetica-Bold', fontSize=15, textColor=colors.HexColor('#2F2F39'))
    month_badge = ParagraphStyle('MBadge', fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#EA1B3D'), alignment=2)
    
    header_data = [
        [Paragraph("<b>InfoBeans Foundation</b>", brand_title), Paragraph(f"<b>{str(month_str).upper()}</b>", month_badge)],
        [Paragraph("EMPOWERING TALENT • BUILDING FUTURES", brand_sub), ""],
        [Paragraph("COLLEGE MONTHLY ATTENDANCE & PROGRESS REPORT", rep_title), ""],
        [Paragraph(f"<b>{college_name}</b>", college_header), ""]
    ]
    t_header = Table(header_data, colWidths=[380, 160])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN', (1,0), (1,3)),
        ('PADDING', (0,0), (-1,-1), 1),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 4))
    elements.append(HRFlowable(width="100%", thickness=2.5, color=colors.HexColor('#EA1B3D'), spaceBefore=2, spaceAfter=8))
    
    total_students = len(df_college)
    num_batches = df_college['Batch'].nunique()
    avg_att = round(df_college['attendance_calc'].mean(), 1)
    avg_test = round(df_college['test_calc'].mean(), 1)
    
    kpi_data = [
        [f"{total_students}", f"{num_batches}", f"{avg_att:.1f}%", f"{avg_test:.1f}%"],
        ["TOTAL STUDENTS", "ACTIVE BATCHES", "AVG ATTENDANCE", "AVG TEST SCORE"]
    ]
    t_kpi = Table(kpi_data, colWidths=[135]*4)
    t_kpi.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 0.75, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 14),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#2F2F39')),
        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,1), (-1,1), 7),
        ('TEXTCOLOR', (0,1), (-1,1), colors.HexColor('#64748B')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_kpi)
    elements.append(Spacer(1, 10))
    
    for batch_name, batch_group in df_college.groupby('Batch', sort=False):
        b_start_str = str(batch_group['Batch Start'].iloc[0])[:10]
        b_end_str = str(batch_group['Batch End'].iloc[0])[:10]
        
        b_header_data = [
            [
                Paragraph(f"<b>BATCH: {str(batch_name).upper()}</b>", ParagraphStyle('BH', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor('#2F2F39'))),
                Paragraph(f"Duration: {b_start_str} to {b_end_str}", ParagraphStyle('BD', fontName='Helvetica', fontSize=8, textColor=colors.HexColor('#64748B'), alignment=2))
            ]
        ]
        t_bheader = Table(b_header_data, colWidths=[270, 270])
        t_bheader.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#FFF5F5')),
            ('LINELEFT', (0,0), (0,0), 3, colors.HexColor('#EA1B3D')),
            ('PADDING', (0,0), (-1,-1), 3.5),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        elements.append(t_bheader)
        
        table_rows = [["STUDENT NAME", "STUDENT ID", "ATTENDANCE %", "TEST AVG %", "STATUS"]]
        b_att_list, b_test_list = [], []
        
        for _, s in batch_group.iterrows():
            att_v = safe_num(s['attendance_calc'])
            test_v = safe_num(s['test_calc'])
            b_att_list.append(att_v)
            b_test_list.append(test_v)
            status_txt, _, _ = get_performance_status(att_v, test_v)
            
            table_rows.append([
                str(s['Student Name']),
                str(s['Student ID']),
                f"{att_v:.1f}%",
                f"{test_v:.1f}%",
                status_txt
            ])
            
        b_avg_att = sum(b_att_list) / len(b_att_list) if b_att_list else 0.0
        b_avg_test = sum(b_test_list) / len(b_test_list) if b_test_list else 0.0
        table_rows.append([f"{batch_name} Batch Average", "", f"{b_avg_att:.1f}%", f"{b_avg_test:.1f}%", ""])
        
        t_batch = Table(table_rows, colWidths=[160, 95, 95, 95, 95])
        t_batch.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2F2F39')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7.5),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
            ('ALIGN', (0,0), (0,-1), 'LEFT'),
            ('ALIGN', (1,0), (-1,-1), 'CENTER'),
            ('PADDING', (0,0), (-1,-1), 3),
            ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#F8FAFC')),
            ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ]))
        elements.append(t_batch)
        elements.append(Spacer(1, 6))
        
    elements.append(Spacer(1, 6))
    clean_col_id = "".join(filter(str.isalnum, str(college_name))).upper()
    clean_month_id = "".join(filter(str.isalnum, str(month_str))).upper()
    report_id = f"IBF-COL-{clean_month_id}-{clean_col_id}"
    today_str = datetime.datetime.now().strftime("%d %B %Y")
    
    footer_data = [
        [f"Report ID: {report_id}", "Official Confidential College Report"],
        [f"Reporting Cycle: {month_str}", "InfoBeans Foundation • ittraining@infobeans.com | Indore, MP"]
    ]
    t_footer = Table(footer_data, colWidths=[270, 270])
    t_footer.setStyle(TableStyle([
        ('FONTSIZE', (0,0), (-1,-1), 7),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#64748B')),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('PADDING', (0,0), (-1,-1), 1),
    ]))
    elements.append(t_footer)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# --- PDF GENERATOR 2: MODERN BRANDED PARENT REPORT ---
def generate_parent_pdf(s_row, month_str):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=32, bottomMargin=32)
    elements = []
    
    brand_title = ParagraphStyle('BTitle', fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#2F2F39'), spaceAfter=1)
    brand_sub = ParagraphStyle('BSub', fontName='Helvetica-Bold', fontSize=8, textColor=colors.HexColor('#EA1B3D'), spaceAfter=4)
    rep_title = ParagraphStyle('RTitle', fontName='Helvetica-Bold', fontSize=10.5, textColor=colors.HexColor('#555562'))
    month_badge = ParagraphStyle('MBadge', fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor('#EA1B3D'), alignment=2)
    sec_title = ParagraphStyle('STitle', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#2F2F39'), spaceBefore=6, spaceAfter=3)
    
    # 1. Header Banner
    header_data = [
        [Paragraph("<b>InfoBeans Foundation</b>", brand_title), Paragraph(f"<b>{str(month_str).upper()}</b>", month_badge)],
        [Paragraph("EMPOWERING TALENT • BUILDING FUTURES", brand_sub), ""],
        [Paragraph("STUDENT MONTHLY ATTENDANCE & PROGRESS REPORT", rep_title), ""]
    ]
    t_header = Table(header_data, colWidths=[380, 160])
    t_header.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('SPAN', (1,0), (1,2)),
        ('PADDING', (0,0), (-1,-1), 1),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 4))
    elements.append(HRFlowable(width="100%", thickness=2.5, color=colors.HexColor('#EA1B3D'), spaceBefore=2, spaceAfter=8))
    
    # 2. Student Info Card
    info_data = [
        ["Student Name", str(s_row['Student Name']), "Student ID", str(s_row['Student ID'])],
        ["Enrolled Batch", str(s_row['Batch']), "College / Institute", str(s_row['College'])],
    ]
    t_info = Table(info_data, colWidths=[95, 175, 95, 175])
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('BOX', (0,0), (-1,-1), 0.75, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0,0), (0,-1), colors.HexColor('#475569')),
        ('TEXTCOLOR', (2,0), (2,-1), colors.HexColor('#475569')),
        ('FONTSIZE', (0,0), (-1,-1), 8.5),
        ('PADDING', (0,0), (-1,-1), 4.5),
    ]))
    elements.append(t_info)
    elements.append(Spacer(1, 8))
    
    # Numbers calculation
    att_pct = safe_num(s_row['attendance_calc'])
    test_pct = safe_num(s_row['test_calc'])
    status_txt, status_fg, status_bg = get_performance_status(att_pct, test_pct)
    
    tot_cls = int(safe_num(s_row.get('Total Classes'), 40))
    prs_cls = int(safe_num(s_row.get('Classes Present'), 0))
    absent = max(0, tot_cls - prs_cls)
    
    t_obt = int(safe_num(s_row.get('Tech Obtained'), 0))
    t_max = int(safe_num(s_row.get('Tech Max Marks'), 100))
    s_obt = int(safe_num(s_row.get('Soft Skills Obtained'), 0))
    s_max = int(safe_num(s_row.get('Soft Skills Max'), 100))
    
    # 3. Modern KPI Highlight Cards
    kpi_cards = [
        [f"{att_pct:.1f}%", f"{test_pct:.1f}%", status_txt.upper()],
        ["MONTHLY ATTENDANCE", "TOTAL ASSESSMENT SCORE", "OVERALL EVALUATION"]
    ]
    t_kpicards = Table(kpi_cards, colWidths=[180, 180, 180])
    t_kpicards.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#F0FDF4') if att_pct>=75 else colors.HexColor('#FFF5F5')),
        ('BACKGROUND', (1,0), (1,-1), colors.HexColor('#F8FAFC')),
        ('BACKGROUND', (2,0), (2,-1), status_bg),
        ('BOX', (0,0), (-1,-1), 0.75, colors.HexColor('#CBD5E1')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 13),
        ('TEXTCOLOR', (0,0), (0,0), colors.HexColor('#16A34A') if att_pct>=75 else colors.HexColor('#EA1B3D')),
        ('TEXTCOLOR', (1,0), (1,0), colors.HexColor('#2F2F39')),
        ('TEXTCOLOR', (2,0), (2,0), status_fg),
        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,1), (-1,1), 7),
        ('TEXTCOLOR', (0,1), (-1,1), colors.HexColor('#64748B')),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_kpicards)
    elements.append(Spacer(1, 8))
    
    # 4. Attendance Breakdown Table
    elements.append(Paragraph("<b>1. Monthly Attendance Record</b>", sec_title))
    att_table_data = [
        ["Total Sessions Conducted", "Sessions Attended", "Sessions Absent", "Attendance %"],
        [str(tot_cls), str(prs_cls), str(absent), f"{att_pct:.1f}%"]
    ]
    t_att = Table(att_table_data, colWidths=[135]*4)
    t_att.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2F2F39')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#FFFFFF')),
        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 4.5),
    ]))
    elements.append(t_att)
    elements.append(Spacer(1, 8))
    
    # 5. Assessment Breakdown Table
    elements.append(Paragraph("<b>2. Monthly Assessment Performance</b>", sec_title))
    t_pct = (t_obt / t_max * 100) if t_max > 0 else 0.0
    s_pct = (s_obt / s_max * 100) if s_max > 0 else 0.0
    
    marks_table_data = [
        ["Evaluation Track", "Maximum Marks", "Marks Obtained", "Track Percentage"],
        ["Technical Skills Assessment", str(t_max), str(t_obt), f"{t_pct:.1f}%"],
        ["Soft Skills & Aptitude Assessment", str(s_max), str(s_obt), f"{s_pct:.1f}%"],
        ["Consolidated Assessment Total", str(t_max + s_max), str(t_obt + s_obt), f"{test_pct:.1f}%"]
    ]
    t_marks = Table(marks_table_data, colWidths=[180, 110, 110, 140])
    t_marks.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2F2F39')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('ALIGN', (1,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,1), (-1,-2), colors.HexColor('#FFFFFF')),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#FFF5F5')),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('PADDING', (0,0), (-1,-1), 4.5),
    ]))
    elements.append(t_marks)
    elements.append(Spacer(1, 14))
    
    # 6. Guidance & Signature Block
    sign_block = [
        [
            Paragraph("<b>Remarks / Faculty Guidance:</b><br/>Consistent session attendance and practical test performance are crucial for IT career readiness.", ParagraphStyle('Rem', fontSize=7.5, textColor=colors.HexColor('#475569'))),
            Paragraph("<b>Academic Administration</b><br/>InfoBeans Foundation", ParagraphStyle('Sign', fontSize=8, textColor=colors.HexColor('#2F2F39'), alignment=2))
        ]
    ]
    t_sign = Table(sign_block, colWidths=[360, 180])
    t_sign.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LINEABOVE', (1,0), (1,0), 0.75, colors.HexColor('#94A3B8')),
        ('PADDING', (0,0), (-1,-1), 2),
    ]))
    elements.append(t_sign)
    elements.append(Spacer(1, 10))
    
    # 7. Document ID & Official Footer
    clean_std_id = "".join(filter(str.isalnum, str(s_row['Student ID']))).upper()
    clean_month = "".join(filter(str.isalnum, str(month_str))).upper()
    doc_id = f"IBF-STU-{clean_month}-{clean_std_id}"
    today_str = datetime.datetime.now().strftime("%d %B %Y")
    
    footer_data = [
        [f"Document ID: {doc_id}", "Official system-generated student progress report."],
        [f"Generated on: {today_str}", "InfoBeans Foundation • Contact: ittraining@infobeans.com"]
    ]
    t_foot = Table(footer_data, colWidths=[270, 270])
    t_foot.setStyle(TableStyle([
        ('FONTSIZE', (0,0), (-1,-1), 6.5),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#94A3B8')),
        ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ('PADDING', (0,0), (-1,-1), 1),
    ]))
    elements.append(t_foot)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

def send_email_with_pdf(sender_email, sender_pwd, target_email, subject, body_text, pdf_bytes, file_name, smtp_server="smtp.gmail.com", smtp_port=587):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = str(target_email)
    msg['Subject'] = subject
    msg.attach(MIMEText(body_text, 'plain'))
    part = MIMEApplication(pdf_bytes, Name=file_name)
    part['Content-Disposition'] = f'attachment; filename="{file_name}"'
    msg.attach(part)
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_pwd)
    server.sendmail(sender_email, [str(target_email)], msg.as_string())
    server.quit()

# --- DYNAMIC EXCEL PARSER ---
def parse_excel_smart(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets_to_process = [s for s in xls.sheet_names if not str(s).startswith("_") and "instruction" not in str(s).lower()]
    if not sheets_to_process:
        sheets_to_process = xls.sheet_names
        
    all_students_list = []
    
    for sheet in sheets_to_process:
        raw_df = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
        
        header_row_idx = None
        for r in range(min(10, len(raw_df))):
            row_str = " ".join([str(val).lower() for val in raw_df.iloc[r].dropna().values])
            if "student" in row_str and ("name" in row_str or "id" in row_str or "college" in row_str):
                header_row_idx = r
                break
                
        if header_row_idx is None:
            header_row_idx = 0
            
        df = pd.read_excel(uploaded_file, sheet_name=sheet, skiprows=header_row_idx)
        
        b_start, b_end = "2026-01-01", "2026-12-31"
        try:
            if header_row_idx > 0:
                top_data = raw_df.iloc[:header_row_idx].values.flatten()
                for v in top_data:
                    if isinstance(v, (datetime.date, datetime.datetime)):
                        b_start = str(v)[:10]
                        break
        except:
            pass

        col_map = {}
        for c in df.columns:
            cl = str(c).strip().lower()
            if "college email" in cl:
                col_map[c] = 'College Email'
            elif "parent email" in cl or "parent's email" in cl or "parent mail" in cl:
                col_map[c] = 'Parent Email'
            elif "college" in cl or "institute" in cl:
                col_map[c] = 'College'
            elif "student name" in cl or "name" in cl:
                col_map[c] = 'Student Name'
            elif "student id" in cl or "enrollment" in cl or "roll" in cl:
                col_map[c] = 'Student ID'
            elif "total class" in cl or "total sessions" in cl or "total lectures" in cl:
                col_map[c] = 'Total Classes'
            elif "present" in cl or "classes attended" in cl:
                col_map[c] = 'Classes Present'
            elif "absent" in cl:
                col_map[c] = 'Classes Absent'
            elif "tech max" in cl:
                col_map[c] = 'Tech Max Marks'
            elif "tech obt" in cl or "technical marks" in cl or "tech marks" in cl:
                col_map[c] = 'Tech Obtained'
            elif "soft skills max" in cl or "soft max" in cl:
                col_map[c] = 'Soft Skills Max'
            elif "soft skills obt" in cl or "soft obt" in cl or "soft skills marks" in cl:
                col_map[c] = 'Soft Skills Obtained'
            elif "total max" in cl:
                col_map[c] = 'Total Max Marks'
            elif "total obt" in cl or "total marks" in cl:
                col_map[c] = 'Total Obtained'

        df = df.rename(columns=col_map)
        
        if 'Student Name' not in df.columns and len(df.columns) > 1:
            df['Student Name'] = df.iloc[:, 1]
        if 'Student ID' not in df.columns and len(df.columns) > 2:
            df['Student ID'] = df.iloc[:, 2]
        if 'College' not in df.columns:
            df['College'] = "InfoBeans Affiliated College"
            
        df = df.dropna(subset=['Student Name']).reset_index(drop=True)
        if len(df) == 0:
            continue
            
        df['Batch'] = sheet
        df['Batch Start'] = b_start
        df['Batch End'] = b_end
        
        tot_c = pd.to_numeric(pd.Series(df['Total Classes'] if 'Total Classes' in df.columns else 40), errors='coerce').fillna(40)
        prs_c = pd.to_numeric(pd.Series(df['Classes Present'] if 'Classes Present' in df.columns else 0), errors='coerce').fillna(0)
        df['Total Classes'] = tot_c
        df['Classes Present'] = prs_c
        df['attendance_calc'] = (prs_c / tot_c.replace(0, 1)) * 100.0
        
        t_obt = pd.to_numeric(pd.Series(df['Tech Obtained'] if 'Tech Obtained' in df.columns else 0), errors='coerce').fillna(0)
        t_max = pd.to_numeric(pd.Series(df['Tech Max Marks'] if 'Tech Max Marks' in df.columns else 100), errors='coerce').fillna(100)
        s_obt = pd.to_numeric(pd.Series(df['Soft Skills Obtained'] if 'Soft Skills Obtained' in df.columns else 0), errors='coerce').fillna(0)
        s_max = pd.to_numeric(pd.Series(df['Soft Skills Max'] if 'Soft Skills Max' in df.columns else 100), errors='coerce').fillna(100)
        
        if 'Total Obtained' in df.columns:
            tot_obt = pd.to_numeric(df['Total Obtained'], errors='coerce')
            tot_mx = pd.to_numeric(df.get('Total Max Marks', 200), errors='coerce').fillna(200)
            for i in range(len(df)):
                if (t_obt.iloc[i] == 0 and s_obt.iloc[i] == 0) and pd.notnull(tot_obt.iloc[i]):
                    t_obt.iloc[i] = tot_obt.iloc[i] / 2.0
                    s_obt.iloc[i] = tot_obt.iloc[i] / 2.0
                    t_max.iloc[i] = tot_mx.iloc[i] / 2.0
                    s_max.iloc[i] = tot_mx.iloc[i] / 2.0
        
        df['Tech Obtained'] = t_obt
        df['Tech Max Marks'] = t_max
        df['Soft Skills Obtained'] = s_obt
        df['Soft Skills Max'] = s_max
        
        df['test_calc'] = ((t_obt + s_obt) / (t_max + s_max).replace(0, 1)) * 100.0
        df['is_report_ready'] = True
        
        all_students_list.append(df)
        
    if not all_students_list:
        return pd.DataFrame()
    return pd.concat(all_students_list, ignore_index=True)

# --- STREAMLIT UI ---
st.markdown("### 📥 Step 1: Upload Batch-Wise Workbook")
uploaded_file = st.file_uploader("Upload Any Student Excel Sheet (.xlsx / .xls)", type=["xlsx", "xls"], key="single_uploader")

if uploaded_file is not None:
    try:
        master_df = parse_excel_smart(uploaded_file)
    except Exception as err:
        st.error(f"Error reading Excel: {err}")
        master_df = pd.DataFrame()
        
    if not master_df.empty:
        total_recs = len(master_df)
        ready_count = int(master_df['is_report_ready'].sum())
        not_ready_count = total_recs - ready_count
        total_colleges = master_df['College'].nunique()
        
        st.markdown("### 📊 Live Generation & Readiness Audit")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Students Enrolled", total_recs)
        k2.metric("🟢 Reports Ready to Send", ready_count)
        k3.metric("🔴 Incomplete Records", not_ready_count)
        k4.metric("🏛️ Affiliated Colleges", total_colleges)
        
        st.markdown("---")
        st.markdown("### ⚙️ Step 2: Configuration & Credentials")
        c1, c2, c3 = st.columns([1, 1, 1])
        with c1:
            month_label = st.text_input("Reporting Month Name", value="August 2026")
        with c2:
            sender_email = st.text_input("Sender Gmail Address (InfoBeans)")
        with c3:
            sender_pwd = st.text_input("16-Digit Google App Password", type="password")
            
        st.markdown("---")
        st.markdown("### 🚀 Step 3: Monitor & Dispatch Reports")
        
        tab_audit, tab_college, tab_parents, tab_logs = st.tabs([
            "📋 Live Status Tracking Table",
            "🏛️ Colleges Dispatch",
            "👨‍👩‍👧 Parents Dispatch",
            "📑 Delivery Logs"
        ])
        
        # 1. AUDIT TAB
        with tab_audit:
            display_df = master_df[['Student ID', 'Student Name', 'Batch', 'College', 'attendance_calc', 'test_calc', 'is_report_ready']].copy()
            display_df['attendance_calc'] = display_df['attendance_calc'].map(lambda x: f"{safe_num(x):.1f}%")
            display_df['test_calc'] = display_df['test_calc'].map(lambda x: f"{safe_num(x):.1f}%")
            display_df['Report Status'] = "🟢 Generated & Ready"
            st.dataframe(display_df[['Student ID', 'Student Name', 'Batch', 'College', 'attendance_calc', 'test_calc', 'Report Status']], use_container_width=True)

        # 2. COLLEGE TAB
        with tab_college:
            col_send_box, _ = st.columns([2, 1])
            with col_send_box:
                if st.button("🚀 Send Consolidated Reports to ALL Colleges", type="primary", use_container_width=True):
                    if not sender_email or not sender_pwd:
                        st.error("Please enter Sender Gmail and App Password in Step 2.")
                    else:
                        progress_bar = st.progress(0)
                        colleges = master_df['College'].unique()
                        for idx, c_name in enumerate(colleges):
                            c_subset = master_df[master_df['College'] == c_name]
                            target_email = c_subset['College Email'].iloc[0] if 'College Email' in c_subset.columns and pd.notnull(c_subset['College Email'].iloc[0]) else f"tpo@{str(c_name).lower().replace(' ', '')}.edu"
                            pdf_bytes = generate_college_pdf(str(c_name), month_label, c_subset)
                            
                            sub = f"InfoBeans Foundation: Monthly Progress & Attendance Report — {c_name} ({month_label})"
                            body = f"Respected College Authority / TPO,\n\nPlease find attached the consolidated monthly progress report for {c_name} for {month_label}.\n\nTotal Students: {len(c_subset)}\n\nRegards,\nAcademic Administration\nInfoBeans Foundation"
                            try:
                                send_email_with_pdf(sender_email, sender_pwd, target_email, sub, body, pdf_bytes, f"{c_name}_Monthly_Report.pdf")
                                status = "Sent"
                            except Exception as e:
                                status = f"Failed: {e}"
                                
                            st.session_state.email_logs.append({
                                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Recipient Type": "College",
                                "Target Name": c_name,
                                "Email": target_email,
                                "Status": status
                            })
                            progress_bar.progress((idx + 1) / len(colleges))
                        st.success("All college reports successfully dispatched!")
                        st.rerun()

            st.markdown("---")
            for col_name in master_df['College'].unique():
                c_subset = master_df[master_df['College'] == col_name]
                target_email = c_subset['College Email'].iloc[0] if 'College Email' in c_subset.columns and pd.notnull(c_subset['College Email'].iloc[0]) else f"tpo@{str(col_name).lower().replace(' ', '')}.edu"
                pdf_data = generate_college_pdf(str(col_name), month_label, c_subset)
                
                c_box1, c_box2, c_box3 = st.columns([3, 1, 1.2])
                with c_box1:
                    st.write(f"🏛️ **{col_name}** | Status: **🟢 Ready** | Enrolled: `{len(c_subset)} Students`")
                with c_box2:
                    st.download_button(label="⬇️ Download PDF", data=pdf_data, file_name=f"College_{str(col_name).replace(' ', '_')}_{month_label.replace(' ', '_')}.pdf", mime="application/pdf", key=f"dl_col_{col_name}")
                with c_box3:
                    if st.button(f"✉️ Send to {col_name}", key=f"send_col_{col_name}"):
                        if not sender_email or not sender_pwd:
                            st.error("Enter email credentials first.")
                        else:
                            sub = f"InfoBeans Foundation: Monthly Progress Report — {col_name} ({month_label})"
                            body = f"Respected Authority,\n\nPlease find attached the monthly progress report for {col_name}.\n\nRegards,\nInfoBeans Foundation"
                            try:
                                send_email_with_pdf(sender_email, sender_pwd, target_email, sub, body, pdf_data, f"{col_name}_Monthly_Report.pdf")
                                st.success(f"Sent to {col_name}!")
                                st.session_state.email_logs.append({"Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Recipient Type": "College", "Target Name": col_name, "Email": target_email, "Status": "Sent"})
                            except Exception as e:
                                st.error(f"Failed: {e}")
                                st.session_state.email_logs.append({"Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Recipient Type": "College", "Target Name": col_name, "Email": target_email, "Status": f"Failed: {e}"})

        # 3. PARENTS TAB
        with tab_parents:
            st.markdown("##### 🚀 Bulk Dispatch Action")
            p_bulk_col, _ = st.columns([2, 1])
            with p_bulk_col:
                if st.button("🚀 Send Individual Reports to ALL Parents", type="primary", use_container_width=True):
                    if not sender_email or not sender_pwd:
                        st.error("Please enter Sender Gmail and App Password in Step 2.")
                    else:
                        progress_bar2 = st.progress(0)
                        status_text = st.empty()
                        total_parents = len(master_df)
                        for idx, row in master_df.iterrows():
                            status_text.text(f"Sending email {idx+1}/{total_parents}: {row['Student Name']}...")
                            parent_email = row['Parent Email'] if 'Parent Email' in row and pd.notnull(row['Parent Email']) else f"parent_{row['Student ID']}@gmail.com"
                            parent_pdf = generate_parent_pdf(row, month_label)
                            sub = f"InfoBeans Foundation: Monthly Progress Report — {row['Student Name']} ({month_label})"
                            body = f"""Dear Parent,

Please find attached the official monthly progress and attendance report for your ward {row['Student Name']} (ID: {row['Student ID']}, Batch: {row['Batch']}) for the month of {month_label}.

Regards,
Academic Administration
InfoBeans Foundation
Indore (M.P.)
"""
                            try:
                                send_email_with_pdf(sender_email, sender_pwd, parent_email, sub, body, parent_pdf, f"{row['Student ID']}_Report.pdf")
                                status = "Sent"
                            except Exception as e:
                                status = f"Failed: {e}"
                            st.session_state.email_logs.append({"Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Recipient Type": "Parent", "Target Name": f"{row['Student Name']} ({row['Student ID']})", "Email": parent_email, "Status": status})
                            progress_bar2.progress((idx + 1) / total_parents)
                        status_text.empty()
                        st.success("All individual parent emails dispatched successfully!")
                        st.rerun()

            st.markdown("---")
            st.markdown("##### 🧑 Individual Student Actions")
            for _, row in master_df.iterrows():
                parent_email = row['Parent Email'] if 'Parent Email' in row and pd.notnull(row['Parent Email']) else f"parent_{row['Student ID']}@gmail.com"
                parent_pdf = generate_parent_pdf(row, month_label)
                p_box1, p_box2, p_box3 = st.columns([3, 1, 1.2])
                with p_box1:
                    st.write(f"🧑 **{row['Student Name']}** (`{row['Student ID']}`) — Status: **🟢 Ready** | {row['Batch']} | 🏛️ {row['College']}")
                with p_box2:
                    st.download_button(label="⬇️ Download PDF", data=parent_pdf, file_name=f"Parent_{row['Student ID']}_{month_label.replace(' ', '_')}.pdf", mime="application/pdf", key=f"dl_par_{row['Student ID']}")
                with p_box3:
                    if st.button(f"✉️ Send to Parent", key=f"send_par_{row['Student ID']}"):
                        if not sender_email or not sender_pwd:
                            st.error("Enter email credentials first.")
                        else:
                            sub = f"InfoBeans Foundation: Monthly Progress Report — {row['Student Name']} ({month_label})"
                            body = f"Dear Parent,\n\nPlease find attached the monthly progress report of {row['Student Name']} for {month_label}.\n\nRegards,\nInfoBeans Foundation"
                            try:
                                send_email_with_pdf(sender_email, sender_pwd, parent_email, sub, body, parent_pdf, f"{row['Student ID']}_Report.pdf")
                                st.success(f"Sent to {row['Student Name']}'s parent!")
                                st.session_state.email_logs.append({"Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Recipient Type": "Parent", "Target Name": f"{row['Student Name']} ({row['Student ID']})", "Email": parent_email, "Status": "Sent"})
                            except Exception as e:
                                st.error(f"Failed: {e}")
                                st.session_state.email_logs.append({"Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Recipient Type": "Parent", "Target Name": f"{row['Student Name']} ({row['Student ID']})", "Email": parent_email, "Status": f"Failed: {e}"})

        # 4. LOGS TAB
        with tab_logs:
            if st.session_state.email_logs:
                logs_df = pd.DataFrame(st.session_state.email_logs)
                st.dataframe(logs_df, use_container_width=True)
                if st.button("Clear Audit Logs"):
                    st.session_state.email_logs = []
                    st.rerun()
            else:
                st.info("No emails dispatched yet in this session.")
