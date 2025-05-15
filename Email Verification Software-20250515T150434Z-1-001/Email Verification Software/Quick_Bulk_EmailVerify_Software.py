# import re
# import smtplib
# import dns.resolver
# import pandas as pd
# import os
# import threading
# import tempfile
# from time import sleep
# from tkinter import *
# from tkinter import filedialog, messagebox, ttk
# import socket

# # ===================== Quick Verify Logic =====================

# def validate_email_syntax(email):
#     email_regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
#     return bool(re.match(email_regex, email))

# def get_mx_records(domain):
#     try:
#         result = dns.resolver.resolve(domain, 'MX')
#         return str(result[0].exchange)
#     except dns.resolver.NoAnswer:
#         return None
#     except dns.resolver.NXDOMAIN:
#         return None
#     except socket.gaierror:
#         return None
#     except Exception as e:
#         return None

# def check_email_active(email):
#     if not validate_email_syntax(email):
#         return False, "Invalid email syntax."

#     domain = email.split('@')[-1]

#     if not domain:
#         return False, "Domain is empty."

#     if len(domain) > 253:
#         return False, "Domain name is too long."

#     labels = domain.split('.')
#     for label in labels:
#         if len(label) > 63:
#             return False, f"Label '{label}' is too long."

#     sleep(3)

#     mail_server = get_mx_records(domain)
#     if not mail_server:
#         return False, f"Domain '{domain}' does not have MX records."

#     try:
#         server = smtplib.SMTP(mail_server, timeout=10)
#         server.set_debuglevel(0)
#         server.helo()
#         code, message = server.mail("test@example.com")
#         if code == 250:
#             code, message = server.rcpt(email)
#             if code == 250:
#                 server.quit()
#                 return True, f"Email '{email}' is valid and active."
#             else:
#                 server.quit()
#                 return False, f"Email '{email}' is not valid (RCPT TO failed)."
#         else:
#             server.quit()
#             return False, f"Error checking email: {message}"
#     except smtplib.SMTPException as e:
#         return False, f"Error checking email: {e}"

# def quick_verify():
#     email = quick_verify_entry.get()
#     if not email:
#         messagebox.showerror("Error", "Please enter an email address.")
#         return

#     is_active, message = check_email_active(email)
#     status = "Active" if is_active else "Inactive"
#     # quick_verify_status.config(text=f"Status: {status}, Message: {message}")
#     # Log result in Main Menu
#     main_menu_text.insert(END, f"Quick Verify: {email} - Status: {status}, Message: {message}\n")
#     main_menu_text.yview(END)  # Auto-scroll to the latest log

# # ===================== Bulk Verify Logic =====================

# def process_excel(input_file):
#     try:
#         df = pd.read_excel(input_file)
#         mandatory_fields = {'firstName', 'lastname', 'webSites'}
#         if not mandatory_fields.issubset(df.columns):
#             raise ValueError(f"Input file must contain the following mandatory fields: {mandatory_fields}")

#         df = df.fillna({'firstName': 'unknown', 'lastname': '', 'webSites': 'unknown', 
#                         'salutation': '', 'functionName': '', 'functionCode': '', 'linkedinprofile': ''})

#         has_id = '_id' in df.columns
#         results = []
#         for index, row in df.iterrows():
#             _id = str(row['_id']).strip() if has_id else None
#             first_name = str(row['firstName']).strip().lower()
#             last_name = str(row['lastname']).strip().lower()
#             websiteNDdomain = str(row['webSites']).strip().lower()
#             salutation = str(row['salutation']).strip()
#             function_name = str(row['functionName']).strip()
#             function_code = str(row['functionCode']).strip()
#             linkedin_profile = str(row['linkedinprofile']).strip()

#             variations = [f"{first_name}@{websiteNDdomain}"]
#             if last_name:
#                 variations.append(f"{first_name}.{last_name}@{websiteNDdomain}")
#             else:  # Only firstName is available
#                 variations = [f"{first_name}@{websiteNDdomain}"]

#             result = {
#                 'firstName': first_name.title(),
#                 'lastname': last_name.title(),
#                 'webSites': websiteNDdomain,
#                 'salutation': salutation,
#                 'functionName': function_name,
#                 'functionCode': function_code,
#                 'linkedinprofile': linkedin_profile,
#                 **{f"EmailVariation{i+1}": email for i, email in enumerate(variations)}
#             }
#             if has_id:
#                 result['_id'] = _id
#             results.append(result)

#         output_df = pd.DataFrame(results)
#         temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
#         output_df.to_excel(temp_file.name, index=False)
#         return temp_file.name
#     except Exception as e:
#         messagebox.showerror("Error", f"Error in processing Excel: {e}")
#         return None

# def stack_email_variations(input_file, output_file):
#     try:
#         df = pd.read_excel(input_file)
#         stacked_rows = []
#         has_id = '_id' in df.columns

#         for index, row in df.iterrows():
#             _id = row['_id'] if has_id else None
#             first_name = row['firstName']
#             last_name = row['lastname']
#             website = row['webSites']
#             salutation = row['salutation']
#             function_name = row['functionName']
#             function_code = row['functionCode']
#             linkedin_profile = row['linkedinprofile']

#             for i in range(1, 9):
#                 email_variation = row.get(f'EmailVariation{i}')
#                 if email_variation:
#                     stacked_row = {
#                         'firstName': first_name,
#                         'lastname': last_name,
#                         'webSites': website,
#                         'salutation': salutation,
#                         'functionName': function_name,
#                         'functionCode': function_code,
#                         'linkedinprofile': linkedin_profile,
#                         'emails': email_variation,
#                     }
#                     if has_id:
#                         stacked_row['_id'] = _id
#                     stacked_rows.append(stacked_row)

#         stacked_df = pd.DataFrame(stacked_rows)
#         stacked_df.to_excel(output_file, index=False)
#         print(f"Stacked email variations saved to {output_file}")
#     except Exception as e:
#         messagebox.showerror("Error", f"Error in stacking email variations: {e}")

# def run_bulk_verify(input_file, output_dir, final_file_name):
#     with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
#         stacked_output = temp_file.name
        

#     email_variation_file = process_excel(input_file)
#     if email_variation_file:
#         stack_email_variations(email_variation_file, stacked_output)

#         df = pd.read_excel(stacked_output)
#         ids = df['_id'].tolist()
#         emails = df['emails'].tolist()
#         first_names = df['firstName'].tolist()
#         last_names = df['lastname'].tolist()
#         websites = df['webSites'].tolist()
#         salutations = df['salutation'].tolist()
#         function_names = df['functionName'].tolist()
#         function_codes = df['functionCode'].tolist()
#         linkedin_profiles = df['linkedinprofile'].tolist()

#         results = []
#         batch_size = 100
#         batch_files = []

#         for i, (_id, email, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile) in enumerate(zip(ids, emails, first_names, last_names, websites, salutations, function_names, function_codes, linkedin_profiles), 1):
#             progress = (i / len(emails)) * 100
#             bulk_progress['value'] = progress
#             main_menu_text.insert(END, f"Processing email {i}/{len(emails)}: {email}\n")
#             main_menu_text.yview(END)
#             # main_menu_text.insert(END, f"Quick Verify: {emails} - Status: {status}, Message: {message}\n")
#             # main_menu_text.yview(END)  # Auto-scroll to the latest log
#             window.update_idletasks()

#             try:
#                 is_active, message = check_email_active(email)
#                 status = "Active" if is_active else "Inactive"
#                 results.append([_id, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile, email, status, message])
#             except Exception as e:
#                 results.append([_id, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile, email, "Inactive", f"Error: {str(e)}"])

#             if len(results) >= batch_size:
#                 batch_df = pd.DataFrame(results, columns=["_id", "firstName", "lastName", "webSites", "salutation", "functionName", "functionCode", "linkedinprofile", "Email", "Status", "Reason"])
#                 batch_file = os.path.join(output_dir, f"Batch_{len(batch_files) + 1}.xlsx")
#                 batch_df.to_excel(batch_file, index=False)
#                 main_menu_text.insert(END, f"Batch saved: {batch_file}\n")
#                 batch_files.append(batch_file)
#                 results.clear()

#         if results:
#             batch_df = pd.DataFrame(results, columns=["_id", "firstName", "lastName", "webSites", "salutation", "functionName", "functionCode", "linkedinprofile", "Email", "Status", "Reason"])
#             batch_file = os.path.join(output_dir, f"Batch_{len(batch_files) + 1}.xlsx")
#             batch_df.to_excel(batch_file, index=False)
#             main_menu_text.insert(END, f"Final batch saved: {batch_file}\n")
#             batch_files.append(batch_file)

#         all_results = []
#         for batch_file in batch_files:
#             batch_df = pd.read_excel(batch_file)
#             all_results.append(batch_df)

#         final_df = pd.concat(all_results, ignore_index=True)
#         final_output_file = os.path.join(output_dir, f"{final_file_name}.xlsx")
#         final_df.to_excel(final_output_file, index=False)
#         main_menu_text.insert(END, f"Final results saved to {final_output_file}\n")
#         messagebox.showinfo("Success", f"Process completed. Results saved to {final_output_file}")
#     else:
#         messagebox.showerror("Error", "Error in email variation generation. Aborting email verification.")

# def browse_bulk_input():
#     file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
#     bulk_input_var.set(file_path)

# def browse_output_dir():
#     folder_path = filedialog.askdirectory()
#     output_dir_var.set(folder_path)

# def start_bulk_verify():
#     input_file = bulk_input_var.get()
#     output_dir = output_dir_var.get()
#     final_file_name = final_file_name_var.get()

#     if not input_file or not output_dir or not final_file_name:
#         messagebox.showerror("Error", "Please fill in all fields before submitting.")
#         return

#     threading.Thread(target=run_bulk_verify, args=(input_file, output_dir, final_file_name)).start()

# # ===================== GUI Layout =====================

# window = Tk()
# window.title("Qick Bulk Email Verification Tool")
# window.geometry("800x600")

# # Flexbox-like layout using grid
# window.grid_columnconfigure(0, weight=1)
# window.grid_columnconfigure(1, weight=1)
# window.grid_rowconfigure(0, weight=1)
# window.grid_rowconfigure(1, weight=1)
# window.grid_rowconfigure(2, weight=1)

# # Quick Verify Section
# quick_verify_frame = LabelFrame(window, text="Quick Verify", padx=10, pady=10)
# quick_verify_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# quick_verify_label = Label(quick_verify_frame, text="Enter Email:")
# quick_verify_label.pack(pady=5)

# quick_verify_entry = Entry(quick_verify_frame, width=30)
# quick_verify_entry.pack(pady=5)

# quick_verify_button = Button(quick_verify_frame, text="Validate", command=quick_verify)
# quick_verify_button.pack(pady=5)

# quick_verify_status = Label(quick_verify_frame, text="Status: N/A", fg="blue")
# quick_verify_status.pack(pady=5)

# # Bulk Verify Section
# bulk_verify_frame = LabelFrame(window, text="Bulk Verify", padx=10, pady=10)
# bulk_verify_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

# bulk_input_label = Label(bulk_verify_frame, text="Upload Excel File:")
# bulk_input_label.pack(pady=5)

# bulk_input_var = StringVar()
# bulk_input_entry = Entry(bulk_verify_frame, textvariable=bulk_input_var, width=30)
# bulk_input_entry.pack(pady=5)

# browse_bulk_input_button = Button(bulk_verify_frame, text="Browse", command=browse_bulk_input)
# browse_bulk_input_button.pack(pady=5)

# output_dir_label = Label(bulk_verify_frame, text="Save Output Directory:")
# output_dir_label.pack(pady=5)

# output_dir_var = StringVar()
# output_dir_entry = Entry(bulk_verify_frame, textvariable=output_dir_var, width=30)
# output_dir_entry.pack(pady=5)

# browse_output_dir_button = Button(bulk_verify_frame, text="Browse", command=browse_output_dir)
# browse_output_dir_button.pack(pady=5)

# final_file_name_label = Label(bulk_verify_frame, text="Final File Name:")
# final_file_name_label.pack(pady=5)

# final_file_name_var = StringVar()
# final_file_name_entry = Entry(bulk_verify_frame, textvariable=final_file_name_var, width=30)
# final_file_name_entry.pack(pady=5)

# start_bulk_verify_button = Button(bulk_verify_frame, text="Start Bulk Verify", command=start_bulk_verify)
# start_bulk_verify_button.pack(pady=10)

# bulk_progress = ttk.Progressbar(bulk_verify_frame, orient="horizontal", length=300, mode="determinate")
# bulk_progress.pack(pady=10)

# # Main Menu Section
# main_menu_frame = LabelFrame(window, text="Main Menu", padx=10, pady=10)
# main_menu_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

# main_menu_text = Text(main_menu_frame, height=10, width=100)
# main_menu_text.pack(pady=10)

# window.mainloop()















































import re
import smtplib
import dns.resolver
import pandas as pd
import os
import threading
import tempfile
from time import sleep
from tkinter import *
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import socket
from tkinter.font import Font

# ===================== Quick Verify Logic =====================

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

def quick_verify():
    email = quick_verify_entry.get()
    if not email:
        messagebox.showerror("Error", "Please enter an email address.")
        return

    is_active, message = check_email_active(email)
    status = "Active" if is_active else "Inactive"
    main_menu_tree.insert("", "end", values=(email, status, message))
    main_menu_tree.yview_moveto(1)  # Auto-scroll to the latest entry

# ===================== Bulk Verify Logic =====================

def process_excel(input_file):
    try:
        df = pd.read_excel(input_file)
        mandatory_fields = {'firstName', 'lastname', 'webSites'}
        if not mandatory_fields.issubset(df.columns):
            raise ValueError(f"Input file must contain the following mandatory fields: {mandatory_fields}")

        df = df.fillna({'firstName': 'unknown', 'lastname': '', 'webSites': 'unknown', 
                        'salutation': '', 'functionName': '', 'functionCode': '', 'linkedinprofile': ''})

        has_id = '_id' in df.columns
        results = []
        for index, row in df.iterrows():
            _id = str(row['_id']).strip() if has_id else None
            first_name = str(row['firstName']).strip().lower()
            last_name = str(row['lastname']).strip().lower()
            websiteNDdomain = str(row['webSites']).strip().lower()
            salutation = str(row['salutation']).strip()
            function_name = str(row['functionName']).strip()
            function_code = str(row['functionCode']).strip()
            linkedin_profile = str(row['linkedinprofile']).strip()

            variations = [f"{first_name}@{websiteNDdomain}"]
            if last_name:
                variations.append(f"{first_name}.{last_name}@{websiteNDdomain}")
            else:  # Only firstName is available
                variations = [f"{first_name}@{websiteNDdomain}"]

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
                result['_id'] = _id
            results.append(result)

        output_df = pd.DataFrame(results)
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_df.to_excel(temp_file.name, index=False)
        # print(f"Email variations saved to temporary file: {temp_file.name}")
        return temp_file.name
    except Exception as e:
        messagebox.showerror("Error", f"Error in processing Excel for email variations: {e}")
        return None

def stack_email_variations(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
        stacked_rows = []
        has_id = '_id' in df.columns

        for index, row in df.iterrows():
            _id = row['_id'] if has_id else None
            first_name = row['firstName']
            last_name = row['lastname']
            website = row['webSites']
            salutation = row['salutation']
            function_name = row['functionName']
            function_code = row['functionCode']
            linkedin_profile = row['linkedinprofile']

            for i in range(1, 9):
                email_variation = row.get(f'EmailVariation{i}')
                if email_variation:
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
                        stacked_row['_id'] = _id
                    stacked_rows.append(stacked_row)

        stacked_df = pd.DataFrame(stacked_rows)
        stacked_df.to_excel(output_file, index=False)
        print(f"Stacked email variations saved to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error in stacking email variations: {e}")

def run_bulk_verify(input_file, output_dir, final_file_name):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
        stacked_output = temp_file.name
        print(stacked_output)

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
            bulk_progress['value'] = progress
            # main_menu_tree.insert("", "end", values=(email, "Processing...", ""))
            # main_menu_tree.yview_moveto(1)  # Auto-scroll to the latest entry
            window.update_idletasks()

            try:
                is_active, message = check_email_active(email)
                status = "Active" if is_active else "Inactive"
                results.append([_id, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile, email, status, message])
                main_menu_tree.insert("", "end", values=(email, status, message))
            except Exception as e:
                results.append([_id, first_name, last_name, website, salutation, function_name, function_code, linkedin_profile, email, "Inactive", f"Error: {str(e)}"])
                main_menu_tree.insert("", "end", values=(email, "Inactive", f"Error: {str(e)}"))

            if len(results) >= batch_size:
                batch_df = pd.DataFrame(results, columns=["_id", "firstName", "lastName", "webSites", "salutation", "functionName", "functionCode", "linkedinprofile", "Email", "Status", "Reason"])
                batch_file = os.path.join(output_dir, f"Batch_{len(batch_files) + 1}.xlsx")
                batch_df.to_excel(batch_file, index=False)
                print(f"Batch saved: {batch_file}\n")
                batch_files.append(batch_file)
                results.clear()
                
        if results:
            batch_df = pd.DataFrame(results, columns=["_id", "firstName", "lastName", "webSites", "salutation", "functionName", "functionCode", "linkedinprofile", "Email", "Status", "Reason"])
            batch_file = os.path.join(output_dir, f"Batch_{len(batch_files) + 1}.xlsx")
            batch_df.to_excel(batch_file, index=False)
            print(f"Final batch saved: {batch_file}\n")
            batch_files.append(batch_file)

        all_results = []
        for batch_file in batch_files:
            batch_df = pd.read_excel(batch_file)
            all_results.append(batch_df)
            
            
            
        final_df = pd.concat(all_results, ignore_index=False)
        final_output_file = os.path.join(output_dir, f"{final_file_name}.xlsx")
        final_df.to_excel(final_output_file, index=False)
        messagebox.showinfo("Success", f"Process completed. Results saved to {final_output_file}")
    else:
        messagebox.showerror("Error", "Error in email variation generation. Aborting email verification.")

def browse_bulk_input():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    bulk_input_var.set(file_path)

def browse_output_dir():
    folder_path = filedialog.askdirectory()
    output_dir_var.set(folder_path)

def start_bulk_verify():
    input_file = bulk_input_var.get()
    output_dir = output_dir_var.get()
    final_file_name = final_file_name_var.get()

    if not input_file or not output_dir or not final_file_name:
        messagebox.showerror("Error", "Please fill in all fields before submitting.")
        return

    threading.Thread(target=run_bulk_verify, args=(input_file, output_dir, final_file_name)).start()

# ===================== GUI Layout =====================

window = Tk()
window.title("Email Bulk Verification Tool")
window.geometry("960x900")  # Increased window size
window.configure(bg="white")

# Custom Font
lexend_font = Font(family="Lexend", size=10)

# Custom Styles
style = ttk.Style()
style.theme_use("clam")
style.configure("TFrame", background="white")
style.configure("TLabel", background="white", foreground="#333333", font=lexend_font)
style.configure("TButton", background="#28a745", foreground="white", font=lexend_font, borderwidth=0, borderradius=30)
style.map("TButton", background=[("active", "#218838")])
style.configure("TEntry", fieldbackground="#f8f9fa", foreground="#333333", font=lexend_font, borderradius=30)
style.configure("Treeview", background="white", foreground="#333333", fieldbackground="white", font=lexend_font)
style.configure("Treeview.Heading", background="#28a745", foreground="white", font=lexend_font)
style.map("Treeview", background=[("selected", "#218838")])
style.configure("Horizontal.TProgressbar", background="#28a745", troughcolor="#f8f9fa", bordercolor="#28a745", lightcolor="#28a745", darkcolor="#28a745")

# --- Grid Configuration for Resizing ---
window.grid_columnconfigure(0, weight=1)  # Left column (Quick/Bulk Verify)
window.grid_columnconfigure(1, weight=1)  # Right column (Main Menu)
window.grid_rowconfigure(1, weight=1)  # Main Menu row


# Define style for the LabelFrame title
style = ttk.Style()
style.configure("Bold.TLabelframe.Label", font=("Arial", 14, "bold"), foreground="black")

# Quick Verify Section
quick_verify_frame = ttk.LabelFrame(window, text="Quick Verify", padding=10, style="Bold.TLabelframe")
quick_verify_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
quick_verify_frame.grid_columnconfigure(0, weight=1) # Make the contents of the frame expand horizontally


quick_verify_label = ttk.Label(quick_verify_frame, text="Enter Email:")
quick_verify_label.pack(pady=5)

quick_verify_entry = ttk.Entry(quick_verify_frame, width=30)
quick_verify_entry.pack(pady=5)

quick_verify_button = ttk.Button(quick_verify_frame, text="Validate", command=quick_verify)
quick_verify_button.pack(pady=5)

# Define style for the LabelFrame title
style = ttk.Style()
style.configure("Bold.TLabelframe.Label", font=("Arial", 14, "bold"), foreground="black")

# Bulk Verify Section
bulk_verify_frame = ttk.LabelFrame(window, text="Bulk Verify", padding=10, style="Bold.TLabelframe")
bulk_verify_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
bulk_verify_frame.grid_columnconfigure(0, weight=1) # Make the contents of the frame expand horizontally


bulk_input_label = ttk.Label(bulk_verify_frame, text="Upload Excel File:")
bulk_input_label.pack(pady=5)

bulk_input_var = StringVar()
bulk_input_entry = ttk.Entry(bulk_verify_frame, textvariable=bulk_input_var, width=30)
bulk_input_entry.pack(pady=5)

browse_bulk_input_button = ttk.Button(bulk_verify_frame, text="Import", command=browse_bulk_input)
browse_bulk_input_button.pack(pady=5)

output_dir_label = ttk.Label(bulk_verify_frame, text="Save Output Directory:")
output_dir_label.pack(pady=5)

output_dir_var = StringVar()
output_dir_entry = ttk.Entry(bulk_verify_frame, textvariable=output_dir_var, width=30)
output_dir_entry.pack(pady=5)

browse_output_dir_button = ttk.Button(bulk_verify_frame, text="Import", command=browse_output_dir)
browse_output_dir_button.pack(pady=5)

final_file_name_label = ttk.Label(bulk_verify_frame, text="Final File Name:")
final_file_name_label.pack(pady=5)

final_file_name_var = StringVar()
final_file_name_entry = ttk.Entry(bulk_verify_frame, textvariable=final_file_name_var, width=30)
final_file_name_entry.pack(pady=5)

start_bulk_verify_button = ttk.Button(bulk_verify_frame, text="Start Bulk Verify", command=start_bulk_verify)
start_bulk_verify_button.pack(pady=10)

bulk_progress = ttk.Progressbar(bulk_verify_frame, orient="horizontal", length=300, mode="determinate", style="Horizontal.TProgressbar")
bulk_progress.pack(pady=10)

## Main Menu Section
main_menu_frame = ttk.LabelFrame(window, text="Main Menu", padding=10)
main_menu_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
main_menu_frame.grid_rowconfigure(0, weight=1)  # Make the Treeview expand vertically

columns = ("Email", "Status", "Reason")  # ... (add other columns as needed)
main_menu_tree = ttk.Treeview(main_menu_frame, columns=columns, show="headings") # Removed fixed height
for col in columns:
    main_menu_tree.heading(col, text=col)
    main_menu_tree.column(col, width=150, stretch=True) # Make columns stretchable
main_menu_tree.pack(fill="both", expand=True, pady=10)


# # Add this function to display instructions
# def show_instructions():
#     instructions = """
#     **Email Verification Tool - Technical Documentation & User Manual**

# **Introduction**

# The Email Verification Tool is designed to verify email addresses in two ways:

# *   **Quick Verify:** Validates a single email inputted by the user.
# *   **Bulk Verify:** Processes an Excel file with a list of names and domains to generate and verify email variations in batches.

# This document covers the tool's functionality, data format requirements, and expected output formats.

# **1. Quick Verification**

# *   **Functionality**
#     *   The user inputs an email address.
#     *   The tool validates the email syntax.
#     *   The tool checks whether the email domain has valid MX records.
#     *   The tool attempts SMTP validation to determine if the email is active.
#     *   Results are displayed in the UI.

# *   **Example**
#     *   **Input:** `akshay.parmar@kompassindia.com`
#     *   **Output:**

#         | Email                               | Status | Reason                    |
#         | :---------------------------------- | :----- | :------------------------ |
#         | `akshay.parmar@kompassindia.com`    | Active | Email is valid and active.|

# **2. Bulk Verification**

# *   **Functionality**
#     *   The user uploads an Excel file containing the necessary data fields.
#     *   The tool generates email variations based on the input.
#     *   The generated emails undergo syntax validation and MX record checks.
#     *   Results are stored in an output file.
#     *   If the input file contains 1000 records, the tool will process them in batches of 100 records.

# *   **Input Data Format**

#     The input file *must* contain the following mandatory fields:

#     | Field Name  | Description                               |
#     | :---------- | :-----------------------------------------|
#     | `firstName` | First name of the person (e.g., "Akshay") |
#     | `lastName`  | Last name of the person (e.g., "Parmar")  |
#     | `webSites`  | Company domain (e.g., "kompassindia.com") |

#     Optional Fields:

#     | Field Name        | Description                                    |
#     | :-----------------| :--------------------------------------------- |
#     | `salutation`      | Salutation (e.g., "Mr.")                       |
#     | `functionName`    | Job title (e.g., "Sales Director/Manager")     |
#     | `functionCode`    | Code associated with the job function          |
#     | `linkedinProfile` | LinkedIn URL of the person                     |

# *   **Example Input**

#     | firstName | lastName | webSites         | salutation  | functionName                 | functionCode | linkedinProfile                               |
#     | :-------- | :------- | :--------------- | :---------- | :----------------------------| :------------| :-------------------------------------------- |
#     | Akshay    | Parmar   | kompassindia.com | Mr.         | Sales Director/Manager       | 5402         | https://linkedin.com/in/akshayparmar          |
#     | Maqbool   | Ansari   | kompassindia.com | Mr.         | Accountant Director/Manager  | 6102         | https://linkedin.com/in/maqboolansari        |

# *   **Generated Email Variations**

#     | firstName | lastName | webSites         | EmailVariation1                | EmailVariation2                   |
#     | :-------- | :------- | :--------------- | :----------------------------- | :---------------------------------|
#     | Akshay    | Parmar   | kompassindia.com | `akshay@kompassindia.com`      | `akshay.parmar@kompassindia.com`  |
#     | Maqbool   | Ansari   | kompassindia.com | `maqbool@kompassindia.com`     | `maqbool.ansari@kompassindia.com` |

# **3. Output File Format**

# The output file contains verification results for each generated email variation.

# *   **Example Output**

#     | firstName | lastName | webSites         | Email                             | Status   | Reason                                          |
#     | :-------- | :------- | :--------------- | :---------------------------------| :------- | :----------------------------|
#     | Akshay    | Parmar   | kompassindia.com | `akshay@kompassindia.com`         | Active   | Email is valid and active.   |
#     | Akshay    | Parmar   | kompassindia.com | `akshay.parmar@kompassindia.com`  | Inactive | Email not found.             |
#     | Maqbool   | Ansari   | kompassindia.com | `maqbool@kompassindia.com`        | Active   | Email is valid and active.   |
#     | Maqbool   | Ansari   | kompassindia.com | `maqbool.ansari@kompassindia.com` | Inactive | Email rejected by server.    |

# **4. How to Use the Tool**

# *   **Step 1: Quick Verify**
#     1.  Open the application.
#     2.  Enter an email address in the Quick Verify field.
#     3.  Click Validate.
#     4.  The result will appear in the Main Menu section.

# *   **Step 2: Bulk Verify**
#     1.  Click "Import" next to "Upload Excel File" to upload an Excel file.
#     2.  Click "Import" next to "Save Output Directory" to select an output directory where batch files will be saved.
#         *   *Note:* The tool processes 100 records per batch (e.g., a 1000-record file will generate 10 batch files).
#     3.  Enter a final output file name.
#     4.  Click Start Bulk Verify.
#     5.  The tool processes emails and displays progress.
#     6.  Once complete, a confirmation message appears, and the file is saved.

#     """
#     # Create the instruction window
#     popup = tk.Toplevel(window)
#     popup.title("Instructions")
#     popup.geometry("700x500")

#     # Create a frame inside the popup
#     frame = ttk.Frame(popup)
#     frame.pack(fill=tk.BOTH, expand=True)

#     # Create a canvas inside the frame
#     canvas = tk.Canvas(frame)
#     canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

#     # Create vertical and horizontal scrollbars
#     scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
#     scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

#     scrollbar_x = ttk.Scrollbar(popup, orient=tk.HORIZONTAL, command=canvas.xview)
#     scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

#     # Create a frame inside the canvas for the text widget
#     text_frame = ttk.Frame(canvas)
#     canvas.create_window((0, 0), window=text_frame, anchor="nw")

#     # Create a text widget
#     text_widget = tk.Text(text_frame, wrap=tk.WORD, padx=20, pady=30, font=("Arial", 10))
#     text_widget.insert(tk.END, instructions)
#     text_widget.config(state=tk.DISABLED)  # Make text non-editable
#     text_widget.pack(fill=tk.BOTH, expand=True)

#     # Configure scrolling
#     canvas.config(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
#     scrollbar_y.config(command=canvas.yview)
#     scrollbar_x.config(command=canvas.xview)

#     def update_scroll_region(event):
#         """Reset the scroll region to encompass the entire text frame."""
#         canvas.configure(scrollregion=canvas.bbox("all"))

#     text_frame.bind("<Configure>", update_scroll_region)
#     canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))


# # Add the Instructions button to the GUI layout
# instructions_button = ttk.Button(window, text="Instructions", command=show_instructions)
# instructions_button.grid(row=2, column=0, columnspan=2, pady=10)  # Place it below the main menu


def show_instructions():
    # Create the instruction window
    popup = tk.Toplevel(window)
    popup.title("Instructions")
    popup.geometry("900x600")  # Increased window size

    # Create a frame inside the popup
    frame = ttk.Frame(popup)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Create a notebook (Tab Layout)
    notebook = ttk.Notebook(frame)
    notebook.pack(fill=tk.BOTH, expand=True)

    # Create frames for different sections
    tab1 = ttk.Frame(notebook)
    tab2 = ttk.Frame(notebook)
    tab3 = ttk.Frame(notebook)
    tab4 = ttk.Frame(notebook)

    notebook.add(tab1, text="Quick Verification")
    notebook.add(tab2, text="Bulk Verification")
    notebook.add(tab3, text="Input & Output Format")
    notebook.add(tab4, text="How to Use")

    # ==========================
    # Quick Verification Tab
    # ==========================
    text_quick = tk.Text(tab1, wrap=tk.WORD, font=("Arial", 10), height=15, width=90)
    quick_instruction = """
    **Quick Verification**

    - The user inputs an email address.
    - The tool validates the email syntax.
    - The tool checks whether the email domain has valid MX records.
    - The tool attempts SMTP validation to determine if the email is active.
    - Results are displayed in the UI.

    **Example:**
    Email: akshay.parmar@kompassindia.com
    Status: Active
    Reason: Email is valid and active.
    """
    text_quick.insert(tk.END, quick_instruction)
    text_quick.config(state=tk.DISABLED)
    text_quick.pack(fill=tk.BOTH, expand=True)

    # ==========================
    # Bulk Verification Tab
    # ==========================
    text_bulk = tk.Text(tab2, wrap=tk.WORD, font=("Arial", 10), height=15, width=90)
    bulk_instruction = """
    **Bulk Verification**

    - The user uploads an Excel file containing names and domains.
    - The tool generates email variations based on the input.
    - The generated emails undergo syntax validation and MX record checks.
    - Results are stored in an output file.
    - Large files are processed in batches of 100 records.

    **Example:**
    - Input File Contains 1000 Records → Processed in 10 Batches.
    """
    text_bulk.insert(tk.END, bulk_instruction)
    text_bulk.config(state=tk.DISABLED)
    text_bulk.pack(fill=tk.BOTH, expand=True)

    # ==========================
    # Input & Output Format Tab (Tables)
    # ==========================
    frame_io = ttk.Frame(tab3)
    frame_io.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # Function to create tables
    def create_table(parent, columns, data):
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=len(data))
        for col in columns:
            tree.heading(col, text=col, anchor="center")
            tree.column(col, anchor="center", stretch=True)
        for row in data:
            tree.insert("", tk.END, values=row)
        tree.pack(fill=tk.BOTH, expand=True, pady=5)
        return tree

    # Mandatory Fields Table
    ttk.Label(frame_io, text="Mandatory Fields", font=("Arial", 10, "bold")).pack()
    create_table(frame_io, ["Field Name", "Description"], [
        ["firstName", "First name (e.g., 'Akshay')"],
        ["lastName", "Last name (e.g., 'Parmar')"],
        ["webSites", "Company domain (e.g., 'kompassindia.com')"]
    ])

    # Optional Fields Table
    ttk.Label(frame_io, text="Optional Fields", font=("Arial", 10, "bold")).pack()
    create_table(frame_io, ["Field Name", "Description"], [
        ["salutation", "Salutation (e.g., 'Mr.')"],
        ["functionName", "Job title (e.g., 'Sales Director')"],
        ["functionCode", "Code associated with the job function"],
        ["linkedinProfile", "LinkedIn URL of the person"]
    ])

    # Example Input Table
    ttk.Label(frame_io, text="Example Input", font=("Arial", 10, "bold")).pack()
    create_table(frame_io, ["firstName", "lastName", "webSites", "salutation", "functionName", "functionCode", "linkedinProfile"], [
        ["Akshay", "Parmar", "kompassindia.com", "Mr.", "Sales Director", "5402", "https://linkedin.com/in/akshayparmar"],
        ["Maqbool", "Ansari", "kompassindia.com", "Mr.", "Accountant", "6102", "https://linkedin.com/in/maqboolansari"]
    ])

    # Generated Email Variations Table
    ttk.Label(frame_io, text="Generated Email Variations", font=("Arial", 10, "bold")).pack()
    create_table(frame_io, ["firstName", "lastName", "webSites", "Email Variation 1", "Email Variation 2"], [
        ["Akshay", "Parmar", "kompassindia.com", "akshay@kompassindia.com", "akshay.parmar@kompassindia.com"],
        ["Maqbool", "Ansari", "kompassindia.com", "maqbool@kompassindia.com", "maqbool.ansari@kompassindia.com"]
    ])

    # Example Output Table
    ttk.Label(frame_io, text="Example Output", font=("Arial", 10, "bold")).pack()
    create_table(frame_io, ["firstName", "lastName", "webSites", "Email", "Status", "Reason"], [
        ["Akshay", "Parmar", "kompassindia.com", "akshay@kompassindia.com", "Active", "Email is valid and active"],
        ["Akshay", "Parmar", "kompassindia.com", "akshay.parmar@kompassindia.com", "Inactive", "Email not found"],
        ["Maqbool", "Ansari", "kompassindia.com", "maqbool@kompassindia.com", "Active", "Email is valid and active"],
        ["Maqbool", "Ansari", "kompassindia.com", "maqbool.ansari@kompassindia.com", "Inactive", "Email rejected by server"]
    ])

    # ==========================
    # How to Use Tab
    # ==========================
    text_how_to_use = tk.Text(tab4, wrap=tk.WORD, font=("Arial", 10), height=15, width=90)
    how_to_use = """
    **Step 1: Quick Verify**
    
    1. Open the application.
    2. Enter an email address in the Quick Verify field.
    3. Click Validate.
    4. The result will appear in the Main Menu section.

    **Step 2: Bulk Verify**
    
    1. Click "Import" next to "Upload Excel File" to upload an Excel file.
    2. Click "Import" next to "Save Output Directory" to select an output directory.
    3. Enter a final output file name with .xlsx.
    4. Click Start Bulk Verify.
    5. The tool processes emails and displays progress.
    6. Once complete, a confirmation message appears, and the file is saved.
    """
    text_how_to_use.insert(tk.END, how_to_use)
    text_how_to_use.config(state=tk.DISABLED)
    text_how_to_use.pack(fill=tk.BOTH, expand=True)

# Add the Instructions button to the GUI layout
instructions_button = ttk.Button(window, text="Instructions", command=show_instructions)
instructions_button.grid(row=2, column=0, columnspan=2, pady=10)  # Place it below the main menu


window.mainloop()
