import configparser
import os
from pathlib import Path
import pandas as pd
import json
import openpyxl
from fpdf import FPDF
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from datetime import datetime

def load_config():
    """Load configuration from config.ini file with environment variable overrides."""
    config = configparser.ConfigParser()
    
    # Load configuration from file
    config_file = Path('config.ini')
    
    # Check if file exists
    if not config_file.exists():
        print(f"Error: Configuration file not found at: {config_file}")
        print("Please create a config.ini file in the same directory as your script.")
        print("\nExample config.ini contents:")
        print("""
[EMAIL]
SMTP_SERVER = smtp.gmail.com
SMTP_PORT = 587
FROM_EMAIL = "demycadwell@gmail.com"
EMAIL_PASSWORD = "eklv plvd gjnv fgot"

[PAYSLIP]
COMPANY_NAME = "revenge fashion"
DEFAULT_STYLE_HEADING = red
DEFAULT_STYLE_BODY = black
PDF_MARGIN_LEFT = 0.75
PDF_MARGIN_RIGHT = 0.75
PDF_MARGIN_TOP = 1.0
PDF_MARGIN_BOTTOM = 0.5
""")
        return None
    
    config.read(config_file)
    
    # Create dictionary with environment variable overrides
    final_config = {}
    
    # EMAIL section
    final_config['SMTP_SERVER'] = os.getenv('SMTP_SERVER', config.get('EMAIL', 'SMTP_SERVER', fallback='smtp.gmail.com'))
    final_config['SMTP_PORT'] = int(os.getenv('SMTP_PORT', config.get('EMAIL', 'SMTP_PORT', fallback='587')))
    final_config['FROM_EMAIL'] = os.getenv('FROM_EMAIL', config.get('EMAIL', 'FROM_EMAIL', fallback='your_email@gmail.com'))
    final_config['EMAIL_PASSWORD'] = os.getenv('EMAIL_PASSWORD', config.get('EMAIL', 'EMAIL_PASSWORD', fallback='your_email_password'))
    
    # PAYSLIP section
    final_config['COMPANY_NAME'] = config.get('PAYSLIP', 'COMPANY_NAME', fallback='Your Company Name')
    final_config['DEFAULT_STYLE_HEADING'] = config.get('PAYSLIP', 'DEFAULT_STYLE_HEADING', fallback='red')
    final_config['DEFAULT_STYLE_BODY'] = config.get('PAYSLIP', 'DEFAULT_STYLE_BODY', fallback='black')
    final_config['PDF_MARGINS'] = {
        'left': float(config.get('PAYSLIP', 'PDF_MARGIN_LEFT', fallback='0.75')),
        'right': float(config.get('PAYSLIP', 'PDF_MARGIN_RIGHT', fallback='0.75')),
        'top': float(config.get('PAYSLIP', 'PDF_MARGIN_TOP', fallback='1.0')),
        'bottom': float(config.get('PAYSLIP', 'PDF_MARGIN_BOTTOM', fallback='0.5'))
    }
    
    return final_config

def create_payslip_directory():
    """Create payslips directory if it doesn't exist."""
    Path('payslips').mkdir(exist_ok=True)

def create_payslip_pdf(employee_data, pdf_path, config):
    """Create a professional payslip PDF for an employee."""
    # Create PDF document
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=config['PDF_MARGINS']['left'] * inch,
        rightMargin=config['PDF_MARGINS']['right'] * inch,
        topMargin=config['PDF_MARGINS']['top'] * inch,
        bottomMargin=config['PDF_MARGINS']['bottom'] * inch
    )
    
    # Initialize elements list
    elements = []
    
    # Create styles with professional color scheme
    styles = {
        'heading': ParagraphStyle(
            name='heading',
            fontSize=16,
            leading=20,
            alignment=1,
            textColor=colors.red,
            fontName='Helvetica-Bold'
        ),
        'subheading': ParagraphStyle(
            name='subheading',
            fontSize=12,
            leading=14,
            alignment=0,
            textColor=colors.black,
            fontName='Helvetica-Bold'
        ),
        'body': ParagraphStyle(
            name='body',
            fontSize=10,
            leading=12,
            alignment=0,
            textColor=colors.black,
            fontName='Helvetica'
        ),
        'signature': ParagraphStyle(
            name='signature',
            fontSize=12,
            leading=14,
            alignment=1,
            textColor=colors.blue,
            fontName='Helvetica-Bold'
        )
    }
    
    # Add header section with red accent
    elements.append(Paragraph(config['COMPANY_NAME'], styles['heading']))
    elements.append(Spacer(1, 0.2 * inch))
    elements.append(Paragraph("Monthly Payslip", styles['subheading']))
    elements.append(Spacer(1, 0.2 * inch))
    
    # Add employee details
    elements.append(Paragraph(f"Employee Name: {employee_data['NAME']}", styles['body']))
    elements.append(Paragraph(f"Employee ID: {employee_data['EMPLOYEE ID']}", styles['body']))
    elements.append(Paragraph(f"Date: {datetime.now().strftime('%B %Y')}", styles['body']))
    elements.append(Spacer(1, 0.3 * inch))
    
    # Add salary details table with professional styling
    salary_data = [
        ['Basic Salary:', f"${employee_data['BASIC SALARY']:,.2f}"],
        ['Allowances:', f"${employee_data['ALLOWANCES']:,.2f}"],
        ['Deductions:', f"${employee_data['DEDUCTIONS']:,.2f}"],
        ['Net Salary:', f"${employee_data['Net salary']:,.2f}"]
    ]
    salary_table = Table(salary_data, style=[
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('BACKGROUND', (0,0), (-1,0), colors.blue),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ])
    elements.append(salary_table)
    elements.append(Spacer(1, 0.3 * inch))
    
    # Add footer section
    elements.append(Paragraph(
        f"{config['COMPANY_NAME']}\n"
        "This payslip was generated automatically.\n"
        "Please contact the payroll department if you notice any discrepancies.",
        styles['body']
    ))
    elements.append(Spacer(1, 0.5 * inch))
    
    # Add professional signature section
    signature_text = """_______________________________
Ms. Jane Smith
Payroll Manager
REVENUE FASHION ZW👒"""
    elements.append(Paragraph(signature_text, styles['signature']))
    
    # Build PDF document
    doc.build(elements)

def generate_payslips(df, config):
    """Generate payslips for all employees in the DataFrame."""
    create_payslip_directory()
    for _, row in df.iterrows():
        try:
            pdf_path = f"payslips/{row['EMPLOYEE ID']}.pdf"
            create_payslip_pdf(row, pdf_path, config)
            print(f"Payslip generated for {row['NAME']} (ID: {row['EMPLOYEE ID']})")
        except Exception as e:
            print(f"Error generating payslip for {row['NAME']}: {str(e)}")

def send_payslip_email(config, employee_data, pdf_path):
    """Send payslip as email attachment to employee."""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['Subject'] = f"Your Monthly Payslip from {config['COMPANY_NAME']}"
        msg['From'] = config['FROM_EMAIL']
        msg['To'] = employee_data['EMAIL']
        
        # Email body
        body = f"""Dear {employee_data['NAME']},
Please find your monthly payslip attached to this email.
Best regards,
{config['COMPANY_NAME']} Payroll Department"""
        
        # Attach message body
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach PDF
        with open(pdf_path, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='pdf')
            attachment.add_header('Content-Disposition', 'attachment',
                                filename=f"{employee_data['EMPLOYEE ID']}_payslip.pdf")
            msg.attach(attachment)
        
        # Send email using TLS
        server = smtplib.SMTP(config['SMTP_SERVER'], config['SMTP_PORT'])
        server.starttls()
        server.login(config['FROM_EMAIL'], config['EMAIL_PASSWORD'])
        server.send_message(msg)
        print(f"Email sent successfully to {employee_data['EMAIL']}")
        server.quit()
        return True
    except smtplib.SMTPAuthenticationError:
        print(f"Email authentication failed for {employee_data['EMAIL']}")
        return False
    except smtplib.SMTPException as e:
        print(f"SMTP error sending email to {employee_data['EMAIL']}: {str(e)}")
        return False
    except Exception as e:
        print(f"Error sending email to {employee_data['EMAIL']}: {str(e)}")
        return False

def send_all_payslips(df, config):
    """Send payslips to all employees in the DataFrame."""
    success_count = 0
    total_employees = len(df)
    
    # First verify all PDFs exist
    missing_pdfs = []
    for _, row in df.iterrows():
        pdf_path = f"payslips/{row['EMPLOYEE ID']}.pdf"
        if not os.path.exists(pdf_path):
            missing_pdfs.append(row['NAME'])
    
    if missing_pdfs:
        print("\nWarning: Missing PDFs for employees:")
        for employee in missing_pdfs:
            print(f"- {employee}")
        response = input("\nContinue sending available payslips? (yes/no): ")
        if response.lower() != 'yes':
            return
    
    # Send emails
    for _, row in df.iterrows():
        pdf_path = f"payslips/{row['EMPLOYEE ID']}.pdf"
        if os.path.exists(pdf_path):
            if send_payslip_email(config, row, pdf_path):
                success_count += 1
    
    print(f"\nSummary:")
    print(f"Total employees: {total_employees}")
    print(f"Emails sent successfully: {success_count}")
    print(f"Failed emails: {total_employees - success_count}")

if __name__ == "__main__":
    # Load configuration
    config = load_config()
    if config is None:
        exit(1)
    
    # Print configuration values to verify
    print("\nConfiguration loaded successfully:")
    print(f"SMTP Server: {config['SMTP_SERVER']}")
    print(f"Company Name: {config['COMPANY_NAME']}")
    
    # Your employee data
    data = {
        'NAME': ['Anisha Gurure', 'Demy Bingura', 'Nashe Pastor'],
        'EMAIL': ['ruvihshh48@gmail.com', 'demycl@gmail.com', 'nashteettth@gmail.com'],
        'BASIC SALARY': [4500.00, 4500.00, 4500.00],
        'ALLOWANCES': [1000.00, 1000.00, 1000.00],
        'DEDUCTIONS': [500.00, 500.00, 500.00],
        'Net salary': [5000.00, 5000.00, 5000.00]
    }
    
    df = pd.DataFrame(data)
    df['EMPLOYEE ID'] = ['A0001', 'A0002', 'A0003']
    
    send_all_payslips(df, config)