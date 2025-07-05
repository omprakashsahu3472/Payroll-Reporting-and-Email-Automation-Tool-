import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from fpdf import FPDF
import smtplib
from email.message import EmailMessage
import ssl
import os
import json
from pathlib import Path

class PayrollApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BOKARO STEEL PLANT(SAIL) Payroll Report Tool")
        self.root.geometry("1536x864")
        self.root.resizable(False, False)
        self.master_path = tk.StringVar()
        self.changes_path = tk.StringVar()
        self.latest_df = pd.DataFrame()
        self.password_visible = False
        self.actual_password = ""
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6, background="#0078D7", foreground="white")
        self.style.configure("TLabel", font=("Segoe UI", 10))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#0078D7", foreground="white")
        self.credentials_file = Path.home() / ".bsp_payroll_credentials.json"
        self.saved_credentials = self.load_credentials()
        self.setup_ui()
        env_email = os.getenv("EMAIL_USER")
        env_pass = os.getenv("EMAIL_PASS")

        if env_email and env_pass:
            self.sender_entry.insert(0, env_email)
            self.password_entry.insert(0, env_pass)
            self.actual_password = env_pass
        else:
            if self.saved_credentials:
                self.sender_entry.insert(0, self.saved_credentials.get('sender_email', ''))
                self.password_entry.insert(0, self.saved_credentials.get('password', ''))
                self.email_entry.insert(0, self.saved_credentials.get('recipient_email', ''))
                self.actual_password = self.saved_credentials.get('password', '')

    def load_credentials(self):
        """Load saved credentials from file"""
        try:
            if self.credentials_file.exists():
                with open(self.credentials_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading credentials: {e}")
        return {}

    def save_credentials(self):
        """Save credentials to file"""
        try:
            credentials = {
                'sender_email': self.sender_entry.get(),
                'password': self.actual_password,
                'recipient_email': self.email_entry.get()
            }
            with open(self.credentials_file, 'w') as f:
                json.dump(credentials, f)
            return True
        except Exception as e:
            print(f"Error saving credentials: {e}")
            return False

    def toggle_password_visibility(self, event=None):
        """Toggle between showing and hiding the password"""
        if not self.password_visible:
            self.actual_password = self.password_entry.get() 
            self.password_entry.config(show='')
            self.password_entry.delete(0, tk.END)
            self.password_entry.insert(0, self.actual_password)
            self.eye_button.config(text="🔒")
        else:
            self.password_entry.config(show='*')
            self.eye_button.config(text="👁")
        
        self.password_visible = not self.password_visible

    def show_password_help(self, event=None):
        """Show password requirements help"""
        messagebox.showinfo(
            "Password Requirements",
            "For Gmail, you need to use an App Password instead of your regular password:\n\n"
            "1. Go to your Google Account Security\n"
            "2. Enable 2-Step Verification if not already enabled\n"
            "3. Create an App Password for this application\n"
            "4. Use the generated 16-digit code as your password\n\n"
            "Note: Make sure to keep this password secure."
        )

    def setup_ui(self):
        frame = tk.Frame(self.root, bg="#f4faff", bd=2, relief="groove")
        frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        try:
            from PIL import Image, ImageTk
            logo = Image.open("D:\\Sail VT\\Payroll\\images.png")
            logo = logo.resize((50, 50), Image.Resampling.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(logo)
            logo_label = tk.Label(frame, image=self.logo_img, bg="#f4faff")
            logo_label.pack(pady=(10, 0))
        except Exception as e:
            print(f"Logo not loaded: {e}")

        tk.Label(frame, text="BOKARO STEEL PLANT", font=("Segoe UI", 18, "bold"), fg="#003366", bg="#f4faff").pack(pady=(10, 0))
        tk.Label(frame, text="हर किसी की ज़िन्दगी से जुड़ा है सेल !", font=("Segoe UI", 12, "italic"), fg="#0059b3", bg="#f4faff").pack(pady=(0, 10))

        input_frame = tk.LabelFrame(frame, text="Upload CSV Files", bg="#f4faff", font=("Segoe UI", 10, "bold"), fg="#003366")
        input_frame.pack(padx=10, pady=10, fill=tk.X)
        
        label_frame = tk.Frame(input_frame, bg="#f4faff")
        label_frame.grid(row=0, column=0, sticky="e", padx=5, pady=5)
        tk.Label(label_frame, text="Master CSV File", bg="#f4faff", font=("Segoe UI", 10)).pack(side="left")
        tk.Label(label_frame, text="*", fg="red", bg="#f4faff", font=("Segoe UI", 10)).pack(side="left")
        tk.Entry(input_frame, textvariable=self.master_path, width=60).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.load_master).grid(row=0, column=2, padx=5, pady=5)

        label_frame2 = tk.Frame(input_frame, bg="#f4faff")
        label_frame2.grid(row=1, column=0, sticky="e", padx=5, pady=5)
        tk.Label(label_frame2, text="Changes CSV File", bg="#f4faff", font=("Segoe UI", 10)).pack(side="left")
        tk.Label(label_frame2, text=" *", fg="red", bg="#f4faff", font=("Segoe UI", 10)).pack(side="left")
        tk.Entry(input_frame, textvariable=self.changes_path, width=60).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(input_frame, text="Browse", command=self.load_changes).grid(row=1, column=2, padx=5, pady=5)

        dropdown_frame = tk.Frame(frame, bg="#f4faff")
        dropdown_frame.pack(pady=(10, 5))
        self.report_option = tk.StringVar()
        report_list = ["Generate Count Report", "Generate Changes Report", "Generate New Joinee Report"]
        self.report_combobox = ttk.Combobox(dropdown_frame, textvariable=self.report_option, values=report_list, state="readonly", width=35)
        self.report_combobox.pack(side=tk.LEFT, padx=5)
        ttk.Button(dropdown_frame, text="Select", command=self.run_selected_report).pack(side=tk.LEFT, padx=5)

        tree_frame = tk.LabelFrame(frame, text="Report Viewer", bg="#f4faff", font=("Segoe UI", 10, "bold"), fg="#003366")
        tree_frame.pack(padx=10, pady=(10, 0), fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        search_frame = tk.Frame(frame, bg="#f4faff")
        search_frame.pack(pady=(10, 5), anchor="e")
        tk.Label(search_frame, text="Search:", bg="#f4faff").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        tk.Entry(search_frame, textvariable=self.search_var).pack(side=tk.LEFT, padx=5)
        ttk.Button(search_frame, text="Find", command=self.search_report).pack(side=tk.LEFT)

        export_frame = tk.Frame(frame, bg="#f4faff")
        export_frame.pack(pady=(10, 5))
        self.export_option = tk.StringVar()
        export_list = ["Export as CSV", "Export as PDF"]
        self.export_combobox = ttk.Combobox(export_frame, textvariable=self.export_option, values=export_list, state="readonly", width=25)
        self.export_combobox.pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Export", command=self.run_export).pack(side=tk.LEFT, padx=5)

        email_frame = tk.Frame(frame, bg="#f4faff")
        email_frame.pack(pady=(10, 5))
        
        attachment_frame = tk.Frame(frame, bg="#f4faff")
        attachment_frame.pack(pady=(5, 10))
        
        tk.Label(attachment_frame, text="Attach as:", bg="#f4faff").pack(side=tk.LEFT)
        self.attachment_option = tk.StringVar()
        attachment_list = ["CSV only", "PDF only", "Both CSV and PDF"]
        self.attachment_combobox = ttk.Combobox(attachment_frame, textvariable=self.attachment_option, 
                                              values=attachment_list, state="readonly", width=15)
        self.attachment_combobox.pack(side=tk.LEFT, padx=5)
        self.attachment_combobox.current(0)  # Set default to "CSV only"
        
        tk.Label(email_frame, text="Recipient Email:", bg="#f4faff").pack(side=tk.LEFT)
        self.email_entry = tk.Entry(email_frame, width=30)
        self.email_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Label(email_frame, text="Your Gmail:", bg="#f4faff").pack(side=tk.LEFT)
        self.sender_entry = tk.Entry(email_frame, width=30)
        self.sender_entry.pack(side=tk.LEFT, padx=5)
        
        pass_label_frame = tk.Frame(email_frame, bg="#f4faff")
        pass_label_frame.pack(side=tk.LEFT)
        
        self.pass_label = tk.Label(pass_label_frame, text="Password: (?)", bg="#f4faff", cursor="hand2")
        self.pass_label.pack(side=tk.LEFT)
        self.pass_label.bind("<Button-1>", self.show_password_help)
        
        pass_frame = tk.Frame(email_frame, bg="#f4faff")
        pass_frame.pack(side=tk.LEFT)
        
        self.password_entry = tk.Entry(pass_frame, show='*', width=22)
        self.password_entry.pack(side=tk.LEFT)
        
        self.eye_button = tk.Button(
            pass_frame, 
            text="👁", 
            relief=tk.FLAT, 
            bg="white", 
            command=self.toggle_password_visibility
        )
        self.eye_button.pack(side=tk.LEFT, padx=(0,5))
        
        ttk.Button(email_frame, text="Send Email", command=self.send_email).pack(side=tk.LEFT, padx=5)

    def search_report(self):
        query = self.search_var.get().lower()
        if not query:
            return
        for row in self.tree.get_children():
            values = self.tree.item(row, "values")
            if any(query in str(value).lower() for value in values):
                self.tree.selection_set(row)
                self.tree.focus(row)
                self.tree.see(row)
                break

    def load_master(self):
        """Load master payroll file with more flexible file selection"""
        try:
            path = filedialog.askopenfilename(
                title="Select Master Payroll File",
                filetypes=[
                    ("CSV Files", "*.csv"),
                    ("Master CSV Files", "MASTER*.csv"),
                    ("All Files", "*.*")
                ],
                initialdir=os.path.expanduser("~") 
            )
            
            if path:
                filename = os.path.basename(path).upper()
                if "MASTER" not in filename:
                    if not messagebox.askyesno(
                        "Confirm File",
                        "This doesn't appear to be a Master file.\n"
                        "Master files should typically have 'MASTER' in the filename.\n"
                        "Do you want to use this file anyway?"
                    ):
                        return  
                self.master_path.set(path)
                print(f"Master file selected: {path}")  
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load master file:\n{e}")
            print(f"Error loading master file: {e}")

    def load_changes(self):
        """Load changes payroll file with more flexible file selection"""
        try:
            path = filedialog.askopenfilename(
                title="Select Changes Payroll File",
                filetypes=[
                    ("CSV Files", "*.csv"),
                    ("Changes CSV Files", "CHANGES*.csv"),
                    ("All Files", "*.*")
                ],
                initialdir=os.path.expanduser("~")  
            )
            
            if path:  
                filename = os.path.basename(path).upper()
                if "CHANGES" not in filename:
                    if not messagebox.askyesno(
                        "Confirm File",
                        "This doesn't appear to be a Changes file.\n"
                        "Changes files should typically have 'CHANGES' in the filename.\n"
                        "Do you want to use this file anyway?"
                    ):
                        return  
                self.changes_path.set(path)
                print(f"Changes file selected: {path}")  
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load changes file:\n{e}")
            print(f"Error loading changes file: {e}")

    def run_selected_report(self):
        choice = self.report_option.get()
        if choice == "Generate Count Report":
            self.generate_count_report()
        elif choice == "Generate Changes Report":
            self.generate_changes_report()
        elif choice == "Generate New Joinee Report":
            self.generate_new_joinee_report()

    def display_report(self, df):
        self.latest_df = df
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            max_len = max(df[col].astype(str).map(len).max(), len(col))
            pixel_width = min(max_len * 8, 300)
            self.tree.column(col, width=pixel_width, anchor='center', minwidth=220, stretch=True)
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def run_export(self):
        if self.latest_df.empty:
            messagebox.showwarning("Warning", "Please generate a report first.")
            return
        choice = self.export_option.get()
        if choice == "Export as CSV":
            self.export_csv()
        elif choice == "Export as PDF":
            self.export_pdf()

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if path:
            self.latest_df.to_csv(path, index=False)
            messagebox.showinfo("Exported", f"CSV saved at:\n{path}")

    def export_pdf(self):
        try:
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if not path:
                return
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=True, margin=10)
            pdf.set_font("Arial", size=6)
            columns = list(self.latest_df.columns)
            usable_width = 290
            rows = self.latest_df.values.tolist()

            for i in range(0, len(columns), 10):
                subset_cols = columns[i:i+10]
                subset_df = self.latest_df[subset_cols]
                pdf.add_page()
                cell_width = max(25, usable_width / len(subset_cols))
                for col in subset_cols:
                    pdf.cell(cell_width, 6, str(col)[:15], border=1)
                pdf.ln()
                for _, row in subset_df.iterrows():
                    for item in row:
                        text = str(item)
                        text = text[:15] + "..." if len(text) > 18 else text
                        pdf.cell(cell_width, 6, text, border=1)
                    pdf.ln()
            pdf.output(path)
            messagebox.showinfo("Exported", f"PDF saved at:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF:\n{e}")

    def export_pdf_to_file(self, filename):
        """Export the current report to PDF without showing save dialog"""
        try:
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=True, margin=10)
            pdf.set_font("Arial", size=6)
            columns = list(self.latest_df.columns)
            usable_width = 290
            rows = self.latest_df.values.tolist()

            for i in range(0, len(columns), 10):
                subset_cols = columns[i:i+10]
                subset_df = self.latest_df[subset_cols]
                pdf.add_page()
                cell_width = max(25, usable_width / len(subset_cols))
                for col in subset_cols:
                    pdf.cell(cell_width, 6, str(col)[:15], border=1)
                pdf.ln()
                for _, row in subset_df.iterrows():
                    for item in row:
                        text = str(item)
                        text = text[:15] + "..." if len(text) > 18 else text
                        pdf.cell(cell_width, 6, text, border=1)
                    pdf.ln()
            pdf.output(filename)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create PDF: {str(e)}")
            return False

    def send_email(self):
        if self.latest_df.empty:
            messagebox.showwarning("Warning", "Generate and export a report first.")
            return

        recipient_email = self.email_entry.get()
        sender_email = self.sender_entry.get()
        password = self.password_entry.get()
        attachment_format = self.attachment_option.get()

        if not all([sender_email, password, recipient_email]):
            messagebox.showwarning("Missing Info", "All fields are required.")
            return

        if not self.save_credentials():
            messagebox.showwarning("Warning", "Could not save credentials for future use.")


        try:
            report_type = self.report_option.get()
            
            
            if report_type == "Generate Count Report":
                base_name = "Employee_Count_Report"
                email_subject = "Employee Count Analysis Report"
            elif report_type == "Generate Changes Report":
                base_name = "Employee_Changes_Report"
                email_subject = "Employee Changes Report"
            elif report_type == "Generate New Joinee Report":
                base_name = "New_Joinees_Report"
                email_subject = "New Employee Joining Report"
            else:
                base_name = "Payroll_Report"
                email_subject = "Payroll Data Report"

            
            attachments = []
            formats_included = []
            
            if attachment_format in ["CSV only", "Both CSV and PDF"]:
                csv_file = f"{base_name}.csv"
                self.latest_df.to_csv(csv_file, index=False)
                attachments.append(csv_file)
                formats_included.append("CSV")
            
            if attachment_format in ["PDF only", "Both CSV and PDF"]:
                pdf_file = f"{base_name}.pdf"
                if self.export_pdf_to_file(pdf_file):
                    attachments.append(pdf_file)
                    formats_included.append("PDF")
                else:
                    if not attachments:  
                        return

           
            msg = EmailMessage()
            msg['Subject'] = f'BSP Payroll System - {email_subject}'
            msg['From'] = sender_email
            msg['To'] = recipient_email
            
            
            if len(formats_included) == 2:
                body = f"""Dear Recipient,

Please find attached the {email_subject} in both CSV and PDF formats.

Regards,
BSP Payroll Team"""
            else:
                format_type = formats_included[0]
                body = f"""Dear Recipient,

Please find attached the {email_subject} in {format_type} format.

Regards,
BSP Payroll Team"""
            
            msg.set_content(body)

            
            for file in attachments:
                with open(file, 'rb') as f:
                    file_data = f.read()
                    if file.endswith('.pdf'):
                        maintype = 'application'
                        subtype = 'pdf'
                    else:
                        maintype = 'text'
                        subtype = 'csv'
                    msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file)

            
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
                server.login(sender_email, password)
                server.send_message(msg)

            messagebox.showinfo("Success", f"{email_subject} sent successfully to {recipient_email}!")
            
            
            for file in attachments:
                if os.path.exists(file):
                    os.remove(file)
                    
        except Exception as e:
            messagebox.showerror("Email Error", f"Failed to send email: {str(e)}")
            
            for file in attachments:
                if os.path.exists(file):
                    os.remove(file)

    def generate_count_report(self):
        try:
            master_df = pd.read_csv(self.master_path.get())
            changes_df = pd.read_csv(self.changes_path.get())
            merged = pd.merge(master_df, changes_df, on="UNIT_PERNO", suffixes=("_old", "_new"))
            changes = []

            for col in set(master_df.columns).intersection(changes_df.columns):
                if col in ("UNIT_PERNO", "YYYYMM"):
                    continue

                old_col = f"{col}_old"
                new_col = f"{col}_new"

                if old_col in merged.columns and new_col in merged.columns:
                    diff_count = (merged[old_col] != merged[new_col]).sum()
                    if diff_count > 0:
                        changes.append((col, diff_count))

            df = pd.DataFrame(changes, columns=["Column", "Count"])
            self.display_report(df)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate count report:\n{e}")

    def generate_changes_report(self):
        try:
            master_df = pd.read_csv(self.master_path.get())
            changes_df = pd.read_csv(self.changes_path.get())
            merged = pd.merge(master_df, changes_df, on="UNIT_PERNO", suffixes=("_old", "_new"))
            report = []
            for col in set(master_df.columns).intersection(changes_df.columns):
                if col in ("UNIT_PERNO", "YYYYMM"):
                    continue
                changed = merged[merged[f"{col}_old"] != merged[f"{col}_new"]]
                for _, row in changed.iterrows():
                    report.append([row["UNIT_PERNO"], col, row[f"{col}_old"], row[f"{col}_new"]])
            df = pd.DataFrame(report, columns=["UNIT_PERNO", "COLUMN", "OLD_VALUE", "NEW_VALUE"])
            self.display_report(df)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate changes report:\n{e}")

    def generate_new_joinee_report(self):
        try:
            master_df = pd.read_csv(self.master_path.get())
            changes_df = pd.read_csv(self.changes_path.get())
            master_df.columns = master_df.columns.str.strip().str.upper()
            changes_df.columns = changes_df.columns.str.strip().str.upper()
            new_joinees = changes_df[~changes_df["UNIT_PERNO"].isin(master_df["UNIT_PERNO"])]
            selected_cols = ["SAIL_PERNO", "NAME", "DOB", "DOJ_SAIL", "PAN", "BANK_ACNO", "IFSC_CD"]
            existing_cols = [col for col in selected_cols if col in new_joinees.columns]
            new_joinees = new_joinees[existing_cols]
            new_joinees.reset_index(drop=True, inplace=True)
            self.display_report(new_joinees)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate new joinee report:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PayrollApp(root)
    root.mainloop()
    
    
