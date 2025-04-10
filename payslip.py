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
import os
from pathlib import Path

# Read the Excel file
df = pd.read_excel("EMPLOYEE PAYSLIP.xlsx")

# Display current columns to verify names
print("Current columns in the DataFrame:")
print(df.columns)

# Select the required columns (using exact column names from your data)
df = df[['NAME', 'EMAIL', 'BASIC SALARY', 'ALLOWANCES', 'DEDUCTIONS']]

# Add EMPLOYEE ID column with default values
df['EMPLOYEE ID'] = ['A0001', 'A0002', 'A0003']

# Net salary calculation (using exact column names from your data)
df['Net salary'] = df['BASIC SALARY'] + df['ALLOWANCES'] - df['DEDUCTIONS']

# Display results with formatted currency
pd.options.display.float_format = '${:,.2f}'.format
print("\nEmployee Salary Detail:")
print(df[['EMPLOYEE ID', 'NAME', 'BASIC SALARY', 'ALLOWANCES', 'DEDUCTIONS', 'Net salary']])

def create_payslip_directory():
    """Create payslips directory if it doesn't exist."""
    Path('payslips').mkdir(exist_ok=True)

def create_payslip_pdf(employee_data, pdf_path):
    """Create a professional payslip PDF for an employee."""
    # Create PDF document
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=1 * inch,
        bottomMargin=0.5 * inch
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
    elements.append(Paragraph("REVENGE FASHION ZW👒", styles['heading']))
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
        "REVENGE FASHION ZW👒\n"
        "This payslip was generated automatically.\n"
        "Please contact the payroll department if you notice any discrepancies.",
        styles['body']
    ))
    elements.append(Spacer(1, 0.5 * inch))
    
    # Add professional signature section
    signature_text = """
    _______________________________
    Ms. Jane Smith
    Payroll Manager
    REVENGE FASHION ZW👒
    """
    elements.append(Paragraph(signature_text, styles['signature']))
    
    # Build PDF document
    doc.build(elements)

def generate_payslips(df):
    """Generate payslips for all employees in the DataFrame."""
    create_payslip_directory()
    for _, row in df.iterrows():
        try:
            pdf_path = f"payslips/{row['EMPLOYEE ID']}.pdf"
            create_payslip_pdf(row, pdf_path)
            print(f"Payslip generated for {row['NAME']} (ID: {row['EMPLOYEE ID']})")
        except Exception as e:
            print(f"Error generating payslip for {row['NAME']}: {str(e)}")

def send_payslip_email(config, employee_data, pdf_path):
    """Send payslip as email attachment to employee."""
    try:
        # Create message
        msg = MIMEMultipart()
        msg['Subject'] = "Your Monthly Payslip from REVENGE FASHION ZW👒"
        msg['From'] = config['FROM_EMAIL']
        msg['To'] = employee_data['EMAIL']
        
        # Email body
        body = f"""
Dear {employee_data['NAME']},
Please find your monthly payslip attached to this email.
Best regards,
REVENGE FASHION ZW👒 Payroll Department
"""
        
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

def load_config():
    """Load configuration from environment variables with default values."""
    config = {
        'SMTP_SERVER': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
        'SMTP_PORT': int(os.getenv('SMTP_PORT', '587')),
        'FROM_EMAIL': os.getenv('FROM_EMAIL', "demycadwell@gmail.com"),
        'EMAIL_PASSWORD': os.getenv('EMAIL_PASSWORD', "eklv plvd gjnv fgot")
    }
    
    # Verify required configuration
    required_keys = ['SMTP_SERVER', 'SMTP_PORT', 'FROM_EMAIL', 'EMAIL_PASSWORD']
    missing_keys = [key for key in required_keys if not config[key]]
    if missing_keys:
        print(f"Missing required configuration: {', '.join(missing_keys)}")
        print("\nPlease set these environment variables or provide default values in the code.")
        return None
    return config

if __name__ == "__main__":
    # Load configuration
    config = load_config()
    if config is None:
        exit(1)

    # Your employee data
    data = {
        'NAME': ['Anisha Gurure', 'Demy Bingura', 'Nashe Pastor'],
        'EMAIL': ['ruvimbo448@gmail.com', 'demycadwell@gmail.com', 'nashegraphix@gmail.com'],
        'BASIC SALARY': [4500.00, 4500.00, 4500.00],
        'ALLOWANCES': [1000.00, 1000.00, 1000.00],
        'DEDUCTIONS': [500.00, 500.00, 500.00],
        'Net salary': [5000.00, 5000.00, 5000.00]
    }
    
    df = pd.DataFrame(data)
    # Add EMPLOYEE ID column
    df['EMPLOYEE ID'] = ['A0001', 'A0002', 'A0003']
    send_all_payslips(df, config)