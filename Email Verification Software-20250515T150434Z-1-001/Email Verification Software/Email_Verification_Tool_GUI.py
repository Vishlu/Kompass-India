# original files
# email_verification.py
import pandas as pd
import os
import re
import smtplib
import dns.resolver
import socket
from time import sleep
import tempfile
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter.ttk import Progressbar, Style
import threading

# ===================== First Part: Email Variation Generation and Stacking =====================

def generate_email_variations(first_name, last_name, websiteNDdomain):
    """
    Generates different email variations based on the inputs.
    """
    full_domain = f"{websiteNDdomain}"
    variations = [
        f"{first_name}.{last_name}@{full_domain}",
        f"{first_name}@{full_domain}",
    ]
    return variations

def process_excel(input_file):
    """
    Reads input data from an Excel file, generates email variations, and writes the results to a temporary Excel file.
    """
    try:
        df = pd.read_excel(input_file)
        print("Column names in the input file:", df.columns)

        # Check for mandatory fields (excluding `_id`)
        mandatory_fields = {'firstName', 'lastname', 'webSites'}
        if not mandatory_fields.issubset(df.columns):
            raise ValueError(f"Input file must contain the following mandatory fields: {mandatory_fields}")

        # Fill missing mandatory fields with default values
        df = df.fillna({'firstName': 'unknown', 'lastname': '', 'webSites': 'unknown', 
                        'salutation': '', 'functionName': '', 'functionCode': '', 'linkedinprofile': ''})

        # Check if `_id` column exists
        has_id = '_id' in df.columns

        results = []
        for index, row in df.iterrows():
            # Include `_id` only if the column exists
            _id = str(row['_id']).strip() if has_id else None
            first_name = str(row['firstName']).strip().lower()
            last_name = str(row['lastname']).strip().lower()
            websiteNDdomain = str(row['webSites']).strip().lower()
            salutation = str(row['salutation']).strip()
            function_name = str(row['functionName']).strip()
            function_code = str(row['functionCode']).strip()
            linkedin_profile = str(row['linkedinprofile']).strip()

            # Generate email variations based on available data
            if last_name:  # Both firstName and lastName are available
                variations = generate_email_variations(first_name, last_name, websiteNDdomain)
            else:  # Only firstName is available
                variations = [f"{first_name}@{websiteNDdomain}"]

            # Create result row
            result = {
                'firstName': first_name.title(),
                'lastname': last_name.title(),
                'webSites': websiteNDdomain,
                'salutation': salutation,
                'functionName': function_name,
                'functionCode': function_code,
                'linkedinprofile': linkedin_profile,
                **{f"EmailVariation{i+1}": email for i, email in enumerate(variations)}
            }
            if has_id:
                result['_id'] = _id  # Add `_id` if it exists

            results.append(result)

        # Save results to a temporary file
        output_df = pd.DataFrame(results)
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_df.to_excel(temp_file.name, index=False)
        print(f"Email variations saved to temporary file: {temp_file.name}")
        return temp_file.name  # Return the temporary file path
    except Exception as e:
        print(f"Error in processing Excel for email variations: {e}")
        return None


def stack_email_variations(input_file, output_file):
    """
    Stacks email variations into a single column format.
    """
    try:
        df = pd.read_excel(input_file)
        stacked_rows = []

        # Check if `_id` column exists
        has_id = '_id' in df.columns

        for index, row in df.iterrows():
            # Include `_id` only if the column exists
            _id = row['_id'] if has_id else None
            first_name = row['firstName']
            last_name = row['lastname']
            website = row['webSites']
            salutation = row['salutation']
            function_name = row['functionName']
            function_code = row['functionCode']
            linkedin_profile = row['linkedinprofile']

            # Iterate over email variations and add them to the stacked rows
            for i in range(1, 9):  # EmailVariation1 to EmailVariation8
                email_variation = row.get(f'EmailVariation{i}')
                if email_variation:  # Only if the variation exists
                    stacked_row = {
                        'firstName': first_name,
                        'lastname': last_name,
                        'webSites': website,
                        'salutation': salutation,
                        'functionName': function_name,
                        'functionCode': function_code,
                        'linkedinprofile': linkedin_profile,
                        'emails': email_variation,
                    }
                    if has_id:
                        stacked_row['_id'] = _id  # Add `_id` if it exists

                    stacked_rows.append(stacked_row)

        # Save stacked results to output file
        stacked_df = pd.DataFrame(stacked_rows)
        stacked_df.to_excel(output_file, index=False)
        print(f"Stacked email variations saved to {output_file}")
    except Exception as e:
        print(f"Error in stacking email variations: {e}")

# ===================== Second Part: Email Verification =====================

def validate_email_syntax(email):
    email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(email_regex, email))

def get_mx_records(domain):
    try:
        result = dns.resolver.resolve(domain, 'MX')
        return str(result[0].exchange)
    except dns.resolver.NoAnswer:
        return None
    except dns.resolver.NXDOMAIN:
        return None
    except socket.gaierror:
        return None
    except Exception as e:
        return None

def check_email_active(email):
    if not validate_email_syntax(email):
        return False, "Invalid email syntax."

    domain = email.split('@')[-1]

    if not domain:
        return False, "Domain is empty."

    if len(domain) > 253:
        return False, "Domain name is too long."

    labels = domain.split('.')
    for label in labels:
        if len(label) > 63:
            return False, f"Label '{label}' is too long."

    sleep(3)

    mail_server = get_mx_records(domain)
    if not mail_server:
        return False, f"Domain '{domain}' does not have MX records."

    try:
        server = smtplib.SMTP(mail_server, timeout=10)
        server.set_debuglevel(0)
        server.helo()
        code, message = server.mail("test@example.com")
        if code == 250:
            code, message = server.rcpt(email)
            if code == 250:
                server.quit()
                return True, f"Email '{email}' is valid and active."
            else:
                server.quit()
                return False, f"Email '{email}' is not valid (RCPT TO failed)."
        else:
            server.quit()
            return False, f"Error checking email: {message}"
    except smtplib.SMTPException as e:
        return False, f"Error checking email: {e}"

# ===================== GUI Design =====================

def run_email_verification(input_file, output_dir, final_file_name):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        stacked_output = temp_file.name
        email_processing_text.insert(tk.END, f"Temporary stacked output file created: {stacked_output}\n")

    email_variation_file = process_excel(input_file)
    if email_variation_file:
        stack_email_variations(email_variation_file, stacked_output)

        df = pd.read_excel(stacked_output)
        ids = df['_id'].tolist()
        emails = df['emails'].tolist()
        first_names = df['firstName'].tolist()
        last_names = df['lastname'].tolist()
        websites = df['webSites'].tolist()
        salutations = df['salutation'].tolist()
        function_names = df['functionName'].tolist()
        function_codes = df['functionCode'].tolist()
        linkedin_profiles = df['linkedinprofile'].tolist()

        results = []
        batch_size = 100
        batch_files = []

        for i, (_id, email, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile) in enumerate(zip(ids, emails, first_names, last_names, websites, salutations, function_names, function_codes, linkedin_profiles), 1):
            progress = (i / len(emails)) * 100
            progress_bar['value'] = progress
            email_processing_text.insert(tk.END, f"Processing email {i}/{len(emails)} (ID: {_id}): {email}\n")
            email_processing_text.yview(tk.END)
            window.update_idletasks()

            try:
                is_active, message = check_email_active(email)
                status = "Active" if is_active else "Inactive"
                results.append([_id, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile, email, status, message])
            except Exception as e:
                results.append([_id, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile, email, "Inactive", f"Error: {str(e)}"])

            if len(results) >= batch_size:
                batch_df = pd.DataFrame(results, columns=["_id", "firstName", "lastName", "webSites", "salutation", "functionName", "functionCode", "linkedinprofile", "Email", "Status", "Reason"])
                batch_file = os.path.join(output_dir, f"Batch_{len(batch_files) + 1}.xlsx")
                batch_df.to_excel(batch_file, index=False)
                email_processing_text.insert(tk.END, f"Batch saved: {batch_file}\n")
                batch_files.append(batch_file)
                results.clear()

        if results:
            batch_df = pd.DataFrame(results, columns=["_id", "firstName", "lastName", "webSites", "salutation", "functionName", "functionCode", "linkedinprofile", "Email", "Status", "Reason"])
            batch_file = os.path.join(output_dir, f"Batch_{len(batch_files) + 1}.xlsx")
            batch_df.to_excel(batch_file, index=False)
            email_processing_text.insert(tk.END, f"Final batch saved: {batch_file}\n")
            batch_files.append(batch_file)

        all_results = []
        for batch_file in batch_files:
            batch_df = pd.read_excel(batch_file)
            all_results.append(batch_df)

        final_df = pd.concat(all_results, ignore_index=True)
        final_output_file = os.path.join(output_dir, f"{final_file_name}.xlsx")
        final_df.to_excel(final_output_file, index=False)
        email_processing_text.insert(tk.END, f"Final results saved to {final_output_file}\n")

        messagebox.showinfo("Success", f"Process completed. Results saved to {final_output_file}")
    else:
        messagebox.showerror("Error", "Error in email variation generation. Aborting email verification.")

def browse_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    input_file_var.set(file_path)

def browse_output_directory():
    folder_path = filedialog.askdirectory()
    output_dir_var.set(folder_path)

def on_submit():
    input_file = input_file_var.get()
    output_dir = output_dir_var.get()
    final_file_name = final_file_name_var.get()

    if not input_file or not output_dir or not final_file_name:
        messagebox.showerror("Error", "Please fill in all fields before submitting.")
        return

    threading.Thread(target=run_email_verification, args=(input_file, output_dir, final_file_name)).start()

window = tk.Tk()
window.title("Bulk Email Verification")
window.geometry("700x700")
window.configure(bg="#1e1e2f")

# Define styles
style = Style()
style.configure("TButton", padding=6, relief="flat", background="#4CAF50", foreground="white")
style.map("TButton", background=[("active", "#45a049")])

# Input file upload
input_file_var = tk.StringVar()
output_dir_var = tk.StringVar()
final_file_name_var = tk.StringVar()

title_label = tk.Label(window, text="Bulk Email Verification", font=("Arial", 20, "bold"), fg="white", bg="#1e1e2f")
title_label.pack(pady=20)

input_frame = tk.Frame(window, bg="#1e1e2f")
input_frame.pack(pady=10)

upload_button = tk.Button(input_frame, text="Upload Excel File", command=browse_input_file, bg="#4CAF50", fg="white")
upload_button.grid(row=0, column=0, padx=10)
input_path_label = tk.Label(input_frame, textvariable=input_file_var, width=50, anchor="w", bg="white", fg="black")
input_path_label.grid(row=0, column=1)

output_frame = tk.Frame(window, bg="#1e1e2f")
output_frame.pack(pady=10)

output_button = tk.Button(output_frame, text="Select Output Directory", command=browse_output_directory, bg="#4CAF50", fg="white")
output_button.grid(row=0, column=0, padx=10)
output_path_label = tk.Label(output_frame, textvariable=output_dir_var, width=50, anchor="w", bg="white", fg="black")
output_path_label.grid(row=0, column=1)

final_file_label = tk.Label(window, text="Save Final File Name:", font=("Arial", 12), fg="white", bg="#1e1e2f")
final_file_label.pack(pady=5)
final_file_entry = tk.Entry(window, textvariable=final_file_name_var, width=40)
final_file_entry.pack()

submit_button = tk.Button(window, text="Submit", command=on_submit, bg="#4CAF50", fg="white")
submit_button.pack(pady=20)

progress_bar = Progressbar(window, orient="horizontal", length=500, mode="determinate", style="green.Horizontal.TProgressbar")
progress_bar.pack(pady=10)

email_processing_label = tk.Label(window, text="Processing Status:", font=("Arial", 12), fg="white", bg="#1e1e2f")
email_processing_label.pack(pady=5)
email_processing_text = tk.Text(window, height=10, width=70, bg="#2b2b3d", fg="white")
email_processing_text.pack(pady=10)

window.mainloop()





