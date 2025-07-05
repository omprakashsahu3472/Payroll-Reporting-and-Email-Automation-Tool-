
# Payroll Reporting and Email Automation Tool

A Python-based desktop application built for Bokaro Steel Plant (SAIL) to streamline payroll data comparison, reporting, and automated email dispatch. The tool provides a simple GUI to compare master and change payroll files, generate actionable reports, and export or email them in CSV and PDF formats.


## 🖥️ Features

- ✅ **Intuitive Tkinter GUI** – No command-line usage needed
- 📂 **Compare CSV Files** – Upload `master_file.csv` and `changes_file.csv` easily
- 📊 **Reports Supported**:
  - **Count Report**: Count of field-wise changes (excludes 0-counts)
  - **Changes Report**: Shows exact field-wise differences
  - **New Joinee Report**: Lists employees newly added in change file
- 💾 **Export Reports**:
  - Export any report as CSV
  - Export any report as well-formatted PDF (auto-pagination)
- 📧 **Send Email**:
  - Attach reports as CSV, PDF, or both
  - Uses **App Passwords** securely via environment variables or local storage
- 🔍 **Search & Scroll** in the built-in viewer

## ⚙️ Technologies Used

- Python 3.10+
- [Tkinter](https://docs.python.org/3/library/tkinter.html) – GUI framework
- [Pandas](https://pandas.pydata.org/) – Data manipulation
- [FPDF](https://pyfpdf.github.io/fpdf2/) – PDF generation
- [smtplib / EmailMessage](https://docs.python.org/3/library/email.message.html) – Email sending
- [openpyxl](https://openpyxl.readthedocs.io/) – (Optional) Excel support


## Environment Variables

To run this project, you will need to add the following environment variables to your .env file

`EMAIL_USER=youremail@example.com`

`EMAIL_PASS=yourpassword`


## 🚀 Deployment

This is a standalone desktop application — no web server or hosting required.

💻 Run Locally

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-username/payroll-reporting-tool.git
   cd payroll-reporting-tool

2. **Install dependencies:**

`pip install -r requirements.txt`

3. **Set environment variables for email automation:**

`set EMAIL_USER=your_gmail@example.com`
`set EMAIL_PASS=your_app_password`

4. **Run the application:**

`python main.py`


## Contributing

Contributions are always welcome!

`Pull requests are welcome! Please open an issue first to discuss changes.`

## Authors

- Om Prakash
📍 India
💼 B.Tech CSE | Python | Automation | Tkinter GUI

