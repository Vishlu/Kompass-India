# import pandas as pd

# def generate_email_variations(first_name, last_name, websiteNDdomain):
#     """
#     Generates different email variations based on the inputs.
#     """
#     full_domain = f"{websiteNDdomain}"
#     variations = [
        
        
#         f"{first_name}@{full_domain}",
#         f"{first_name}{last_name}@{full_domain}",
#         f"{first_name}.{last_name}@{full_domain}",
#         f"{first_name}{last_name[0]}@{full_domain}",
#         f"{first_name[0]}.{last_name}@{full_domain}",
#     ]
#     return variations

# def process_excel(input_file, output_file):
#     """
#     Reads input data from an Excel file, generates email variations, and writes the results to a new Excel file.
#     """
#     # Read the Excel file
#     df = pd.read_excel(input_file)

#     # Debugging: Print column names to verify
#     print("Column names in the input file:", df.columns)

#     # Ensure the required columns exist
#     if not {'_id', 'firstName', 'lastname', 'webSites'}.issubset(df.columns):
#         raise ValueError("Input file must contain '_id', 'firstName', 'lastname', and 'webSites' columns.")

#     # Fill NaN values with default strings
#     df = df.fillna({'firstName': 'unknown', 'lastname': 'unknown', 'webSites': 'unknown'})

#     results = []
    
#     # Loop through each row in the dataframe
#     for index, row in df.iterrows():
#         _id = str(row['_id']).strip()  # Retaining the _id
#         first_name = str(row['firstName']).strip().lower()
#         last_name = str(row['lastname']).strip().lower()
#         websiteNDdomain = str(row['webSites']).strip().lower()
        
#         # Generate email variations
#         variations = generate_email_variations(first_name, last_name, websiteNDdomain)
        
#         # Append the result to the results list, including the _id
#         results.append({
#             '_id': _id,
#             'firstName': first_name,
#             'lastname': last_name,
#             'webSites': websiteNDdomain,
#             **{f"EmailVariation{i+1}": email for i, email in enumerate(variations)}
#         })

#     # Convert results to a DataFrame and write to the output file
#     output_df = pd.DataFrame(results)
#     output_df.to_excel(output_file, index=False)
#     print(f"Email variations saved to {output_file}")

# def main():
#     """
#     Main function to process the uploaded file.
#     """
#     # Hardcoded file paths
#     input_file = r"C:\Users\Kompass\Downloads\Marketing Manager_SERP Data Extraction_20-01-2025.xlsx"
#     output_file = r"C:\Users\Kompass\Downloads\Multiple_Email_variation_24.01.2025.xlsx"

#     # Process the input file and generate the email variations
#     process_excel(input_file, output_file)

# if __name__ == "__main__":
#     main()













import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def generate_email_variations(first_name, last_name, full_domain):
    
    var_sales = "sales"
    var_marketing = "marketing"
    var_mktg = "mktg"
    var_export = "export"
    variations = [
        f"{first_name}@{full_domain}",
        f"{first_name}{last_name}@{full_domain}",
        # f"{first_name}.{last_name}@{full_domain}",
        
        # f"{first_name[0]}{last_name}@{full_domain}",  # aparmar@kompassindia.com
        # f"{first_name[0]}.{last_name}@{full_domain}", # a.parmar@kompassindia.com
        # f"{first_name}{last_name[0]}@{full_domain}",# akshayp@kompassindia.com
        # f"{first_name}.{last_name[0]}@{full_domain}",# akshay.p@kompassindia.com
        # f"{last_name}@{full_domain}",  #parmar@kompassindia.com
        # f"{last_name}{first_name[0]}@{full_domain}",  #parmara@kompassindia.com
        # f"{var_sales}@{full_domain}",
        f"{var_marketing}@{full_domain}",
        # f"{var_mktg}@{full_domain}",
        # f"{var_export}@{full_domain}",

    ]
    return variations

def process_and_stack_emails(input_file, output_file):
    try:
        df = pd.read_excel(input_file)
        print("Column names in the input file:", df.columns)

        mandatory_fields = {'First Name', 'Last Name', 'Website'}
        if not mandatory_fields.issubset(df.columns):
            raise ValueError(f"Input file must contain the following mandatory fields: {mandatory_fields}")

        df = df.fillna({'First Name': '', 'Last Name': '', 'Website': ''})
        has_id = '_id' in df.columns
        extra_fields = set(df.columns) - mandatory_fields - {'_id'}

        results = []
        for _, row in df.iterrows():
            first_name = str(row['First Name']).strip().lower()
            last_name = str(row['Last Name']).strip().lower()
            websiteNDdomain = str(row['Website']).strip().lower()

            if last_name:
                variations = generate_email_variations(first_name, last_name, websiteNDdomain)
            else:
                variations = [f"{first_name}@{websiteNDdomain}"]

            extra_data = {field: row[field] for field in extra_fields}

            for email in variations:
                result = {
                    'First Name': first_name.title(),
                    'Last Name': last_name.title(),
                    'Website': websiteNDdomain,
                    'Email': email,
                }
                if has_id:
                    result['_id'] = row['_id']
                result.update(extra_data)
                results.append(result)

        output_df = pd.DataFrame(results)
        output_df.to_excel(output_file, index=False)
        print(f"Stacked email variations saved to {output_file}")

    except Exception as e:
        print(f"Error: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

def browse_input_file():
    input_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    input_file_var.set(input_file)

def submit():
    input_file = input_file_var.get()
    file_name = output_file_name_var.get()
    
    if not input_file or not file_name:
        messagebox.showwarning("Warning", "Please select both input file and provide a file name.")
        return

    # Ask user where to save the output file
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=f"{file_name}.xlsx")
    
    if output_file:
        process_and_stack_emails(input_file, output_file)
        messagebox.showinfo("Success", f"Email variations saved to {output_file}")

# Setting up the GUI
root = tk.Tk()
root.title("Email Variation Tool")  # Set title
root.geometry("480x680")
root.configure(bg="white")  # White background

input_file_var = tk.StringVar()
output_file_name_var = tk.StringVar()

# Text style: Bold and using a good font family for normal text
label_font = ("Helvetica", 14, "bold")
entry_font = ("Helvetica", 12)

# Input file selection
input_label = tk.Label(root, text="Email Variation Tool", bg="white", fg="black", font=label_font)  # Black text, bold
input_label.pack(pady=20)
input_button = tk.Button(root, text="Browse", command=browse_input_file, bg="#800000", fg="white", font=("Helvetica", 12), relief="flat", padx=10, pady=5)  # Maroon button
input_button.pack(pady=10)
input_path_label = tk.Label(root, textvariable=input_file_var, bg="white", fg="black", font=("Helvetica", 10))  # Black text
input_path_label.pack(pady=5)

# File name input
file_name_label = tk.Label(root, text="Enter Output File Name", bg="white", fg="black", font=label_font)  # Black text, bold
file_name_label.pack(pady=20)
file_name_entry = tk.Entry(root, textvariable=output_file_name_var, font=entry_font, width=30, relief="solid", bd=1)
file_name_entry.pack(pady=10)

# Submit Button
submit_button = tk.Button(root, text="Submit", command=submit, bg="#800000", fg="white", font=("Helvetica", 12), relief="flat", padx=10, pady=5)  # Maroon button
submit_button.pack(pady=30)

root.mainloop()
