# Payslip Generator

## Description
The Payslip Generator is a Python-based application designed to automate the process of generating and emailing employee payslips. It reads employee salary details from an Excel file, calculates net salaries, creates payslip PDFs, and sends them via email.

## Features
- Reads employee data from an Excel file
- Calculates net salary (Basic Salary + Allowances - Deductions)
- Generates payslip PDFs for each employee
- Sends payslips via email using SMTP
- Uses multithreading for efficient processing

## Technologies Used
- Python
- Pandas (for data processing)
- ReportLab (for PDF generation)
- FPDF (alternative PDF generation library)
- smtplib (for sending emails)
- concurrent.futures (for parallel processing)

## Installation
### Prerequisites
Ensure you have Python installed. You can install required dependencies using:
```sh
pip install pandas fpdf reportlab openpyxl
```

## Usage
1. Place the employee data in an Excel file (`employees.xlsx`) with the following columns:
   - Employee ID
   - Name
   - Email
   - Basic Salary
   - Allowances
   - Deductions
2. Configure SMTP settings in the script:
```python
smtp_config = {
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 587,
    'from_email': 'your-email@gmail.com',
    'password': 'your-email-password'
}
```
3. Run the script:
```sh
python payslip_generator.py
```

## Output
- PDF payslips are generated in the `payslips/` directory.
- Emails are sent with the payslip attached.

## Error Handling
- If the Excel file is missing required columns, an error message is displayed.
- Any issues during email sending or file creation are logged to the console.

## Future Improvements
- Add a GUI interface for user-friendly interaction.
- Improve security by using environment variables for SMTP credentials.
- Support additional file formats for employee data (CSV, JSON).

## License
This project is open-source and available for modification and distribution.

## Author
Nashe Graphix Pvt Ltd

