# Payslip Generation System

A Python-based system for generating professional payslips and distributing them via email.

## Overview

This system automates the process of creating formatted payslips and sending them to employees. It supports:

* Professional PDF generation with customizable layouts
* Automated email distribution
* Configurable company branding
* Error handling and logging
* Environment variable support

## Prerequisites

Before running the system, ensure you have:

### Required Libraries

```bash
pip install pandas openpyxl reportlab fpdf
```

### Configuration Requirements

Create a 
 file in the project root directory:

```ini
[EMAIL]
SMTP_SERVER = smtp.gmail.com
SMTP_PORT = 587
FROM_EMAIL = "your_company_email@gmail.com"
EMAIL_PASSWORD = "your_app_password"

[PAYSLIP]
COMPANY_NAME = "Your Company Name"
DEFAULT_STYLE_HEADING = red
DEFAULT_STYLE_BODY = black
PDF_MARGIN_LEFT = 0.75
PDF_MARGIN_RIGHT = 0.75
PDF_MARGIN_TOP = 1.0
PDF_MARGIN_BOTTOM = 0.5
```

## Setup Instructions

1. Clone the repository:
   ```bash
git clone https://github.com/your-repo/payslip-generator.git
cd payslip-generator
```
2. Install dependencies:
   ```bash
pip install -r requirements.txt
```
3. Create configuration file:
   - Copy the template above to 

   - Update all fields with your organization's details
   - For Gmail users, generate an App Password instead of using your regular password

## Running the System

Prepare your employee data as a CSV or Excel file with the following columns:

| Column Name | Description | Example |
|-------------|-------------|---------|
| NAME        | Employee name | John Doe |
| EMAIL       | Employee email | john@example.com |
| EMPLOYEE ID | Unique identifier | E001 |
| BASIC SALARY | Base salary amount | 4500.00 |
| ALLOWANCES  | Additional allowances | 1000.00 |
| DEDUCTIONS  | Salary deductions | 500.00 |

Run the script:
```bash
python payslip_generator.py
```

## Features

* **Professional PDF Generation**
  - Customizable margins and styling
  - Company branding integration
  - Professional formatting with tables
  - Automated signature section

* **Email Distribution**
  - Secure SMTP authentication
  - Error handling for failed deliveries
  - Summary report generation
  - Attachment management

* **Configuration Options**
  - Environment variable overrides
  - Customizable PDF styling
  - Flexible email settings
  - Directory management

## Environment Variables

Optional environment variables can override config.ini values:

```bash
export SMTP_SERVER="smtp.example.com"
export FROM_EMAIL="hr@example.com"
export EMAIL_PASSWORD="your_app_password"
```

## Troubleshooting

Common issues and solutions:

1. Email Authentication Failed:
   - Verify SMTP credentials
   - Check firewall settings
   - Ensure App Password is used (Gmail)

2. Missing PDFs:
   - Verify employee data completeness
   - Check permissions on payslips directory
   - Review error logs

3. PDF Generation Errors:
   - Check font availability
   - Verify margin settings
   - Ensure reportlab installation

## Contributing

Pull requests welcome! Please include:
- Test cases for new features
- Documentation updates
- Error handling improvements

## License

[MIT License](LICENSE.md)

## Contact

For issues or feature requests, please open a ticket on GitHub.
