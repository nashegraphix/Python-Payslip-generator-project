import os
import pandas as pd
from fpdf import FPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from datetime import datetime
from pathlib import Path
import concurrent.futures
import smtplib
from email.mime.multipart import MIMEMultipart  # Add this import
from email.mime.base import MIMEBase
from email.mime.text import MIMEText  # Add this import
from email.utils import formatdate
from email import encoders

def load_employee_data(file_path):
    """Load employee data from Excel file"""
    try:
        df = pd.read_excel(file_path)
        required_columns = ['Employee ID', 'Name', 'Email', 'Basic Salary', 'Allowances', 'Deductions']
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Excel file missing required columns")
        
        # Calculate Net Salary
        df['Net Salary'] = df['Basic Salary'] + df['Allowances'] - df['Deductions']
        
        # Validate employee data
        for _, row in df.iterrows():
            if pd.isna(row['Employee ID']) or pd.isna(row['Email']):
                raise ValueError(f"Missing data for employee: {row['Name']}")
        
        return df
    except Exception as e:
        print(f"Error loading employee data: {str(e)}")
        return None

def create_payslip(employee_data):
    """Create payslip PDF for a single employee"""
    Path('payslips').mkdir(exist_ok=True)
    
    # Create PDF
    employee_id = str(employee_data['Employee ID'])  # Ensure string format
    filename = f"payslips/{employee_id}.pdf"
    
    c = canvas.Canvas(filename, pagesize=A4)
    
    # Company Header with gradient
    c.setFillColor(colors.purple)
    c.rect(20*mm, 280*mm, 160*mm, 20*mm, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(20*mm, 285*mm, "Nashe Graphix Pvt Ltd")
    
    # Tagline
    c.setFillColor(colors.purple)
    c.setFont("Helvetica", 15)
    c.drawString(20*mm, 270*mm, "MONTHLY PAYSLIP")
    
    # Date
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 10)
    current_date = datetime.now().strftime("%B %d, %Y")
    c.drawString(20*mm, 260*mm, f"Date: {current_date}")
    
    # Employee Details Table
    c.setFont("Helvetica-Bold", 12)
    c.drawString(20*mm, 240*mm, "Employee Details")
    c.line(20*mm, 235*mm, 180*mm, 235*mm)
    
    # Employee Information
    c.setFont("Helvetica", 10)
    c.drawString(20*mm, 220*mm, f"Employee ID: {employee_id}")
    c.drawString(20*mm, 210*mm, f"Name: {employee_data['Name']}")
    c.drawString(20*mm, 200*mm, f"Email: {employee_data['Email']}")
    
    # Salary Breakdown Table
    c.setFont("Helvetica-Bold", 12)
    c.drawString(20*mm, 180*mm, "Salary Breakdown")
    c.line(20*mm, 175*mm, 180*mm, 175*mm)
    
    # Salary Details
    c.setFont("Helvetica", 10)
    c.drawString(20*mm, 160*mm, f"💼 Basic Salary: ${employee_data['Basic Salary']:,.2f}")
    c.drawString(20*mm, 150*mm, f"➕ Allowances: ${employee_data['Allowances']:,.2f}")
    c.drawString(20*mm, 140*mm, f"➖ Deductions: ${employee_data['Deductions']:,.2f}")
    
    # Net Salary Highlight
    c.setFillColor(colors.purple)
    c.setFont("Helvetica-Bold", 14)
    net_salary = employee_data['Net Salary']
    c.drawString(20*mm, 120*mm, f"📊 Net Salary: ${net_salary:,.2f}")
    
    c.save()
    return filename

def send_email(to_email, subject, body, attachment_path, smtp_config):
    """Send email with attachment"""
    try:
        msg = MIMEMultipart()
        msg['From'] = smtp_config['from_email']
        msg['To'] = to_email
        msg['Date'] = formatdate(localtime=True)
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach the payslip
        with open(attachment_path, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 
                          f'attachment; filename= {os.path.basename(attachment_path)}')
            msg.attach(part)
        
        # Connect to SMTP server
        server = smtplib.SMTP(smtp_config['smtp_server'], smtp_config['smtp_port'])
        server.starttls()
        server.login(smtp_config['from_email'], smtp_config['password'])
        
        # Send email
        server.sendmail(smtp_config['from_email'], to_email, msg.as_string())
        server.quit()
        
        print(f"Email sent successfully to {to_email}")
        return True
    except Exception as e:
        print(f"Error sending email to {to_email}: {str(e)}")
        return False

def generate_payslips_batch(file_path):
    """Generate payslips for all employees in batch"""
    df = load_employee_data(file_path)
    if df is None:
        return []
    
    generated_files = []
    
    # Process employees in parallel
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for _, row in df.iterrows():
            employee_data = row.to_dict()
            employee_data['Employee ID'] = str(employee_data['Employee ID'])
            future = executor.submit(create_payslip, employee_data)
            futures.append(future)
        
        for future in concurrent.futures.as_completed(futures):
            try:
                generated_files.append(future.result())
            except Exception as e:
                print(f"Error generating payslip: {str(e)}")
    
    return generated_files

def send_payslips(generated_files, df, smtp_config):
    """Send payslips to all employees"""
    sent_count = 0
    
    # Process emails in parallel
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for file_path, (_, row) in zip(generated_files, df.iterrows()):
            email_body = f"""
Dear {row['Name']},

Please find your payslip for this month attached to this email.

Best regards,
[Your Company Name]
"""
            future = executor.submit(send_email, 
                                   row['Email'],
                                   "Your Payslip for This Month",
                                   email_body,
                                   file_path,
                                   smtp_config)
            futures.append(future)
        
        for future in concurrent.futures.as_completed(futures):
            if future.result():
                sent_count += 1
    
    return sent_count

def main():
    excel_file = "employees.xlsx"
    print("Generating payslips...")
    
    # SMTP configuration
    smtp_config = {
        'smtp_server': 'smtp.gmail.com',
        'smtp_port': 587,
        'from_email': 'nashegraphix@gmail.com',
        'password': 'jrjd eqgl akox kepa'
    }
    
    # Generate payslips
    df = load_employee_data(excel_file)
    if df is None:
        return
    
    generated_files = generate_payslips_batch(excel_file)
    
    # Send emails
    print("\nSending emails...")
    sent_count = send_payslips(generated_files, df, smtp_config)
    
    print("\nEmail sending complete!")
    print(f"Successfully sent {sent_count} emails out of {len(generated_files)}")

if __name__ == "__main__":
    main()