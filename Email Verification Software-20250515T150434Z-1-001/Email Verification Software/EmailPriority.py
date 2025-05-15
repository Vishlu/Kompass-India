# # firstNmae, lastname, webSites


# import pandas as pd

# # Function to determine email priority
# def determine_priority(row):
#     email = row['Email']
#     first_name = row['firstName']
#     last_name = row['lastName']
#     full_domain = row['webSites']
    
#     # Priority P1 for {first_name}.{last_name}@{full_domain}
#     if email == f"{first_name.lower()}.{last_name.lower()}@{full_domain.lower()}":
#         return 'P1'
    
#     # Priority P2 for {first_name}@{full_domain}
#     elif email == f"{first_name.lower()}@{full_domain.lower()}":
#         return 'P2'
    
#     # Eliminate any other variation
#     return None

# # Function to process the Excel data
# def process_excel(file_path, output_path):
#     # Read the Excel file into a DataFrame
#     df = pd.read_excel(file_path)
    
#     # Filter out rows with 'Inactive' status
#     df = df[df['Status'] == 'Active']
    
#     # Apply the logic to determine the priority
#     df['Priority'] = df.apply(determine_priority, axis=1)
    
#     # Filter out rows where the priority is None (i.e., unwanted email variations)
#     df = df[df['Priority'].notna()]
    
#     # Write the processed data to an output file (Excel format)
#     df.to_excel(output_path, index=False)
#     print(f"Processed data saved to: {output_path}")

# # Example usage
# input_file = r"C:\Users\Kompass\Desktop\SCRIPTS_Company Update\WORKING_Scripts_Company Update\output\PriorityInput.xlsx"
# output_file = r"C:\Users\Kompass\Desktop\SCRIPTS_Company Update\WORKING_Scripts_Company Update\output\750_ActiveEmail_PriorityOutput.xlsx"
# process_excel(input_file, output_file)






# import pandas as pd

# # Function to determine email priority
# def determine_priority(email, first_name, last_name, domain):
#     # Priority P1 for {first_name}.{last_name}@{domain}
#     if email == f"{first_name.lower()}.{last_name.lower()}@{domain.lower()}":
#         return 'P1'
#     # Priority P2 for {first_name}@{domain}
#     elif email == f"{first_name.lower()}@{domain.lower()}":
#         return 'P2'
#     # Eliminate other variations
#     return None

# # Function to process the Excel data
# def process_excel(file_path, output_path):
#     # Read the Excel file into a DataFrame
#     df = pd.read_excel(file_path)
    
#     # Initialize a list to store the processed rows
#     processed_rows = []
    
#     # Group by 'webSites' to process each domain separately
#     grouped = df.groupby(['webSites'])

#     for _, group in grouped:
#         # Iterate over each row in the group
#         group['Priority'] = group.apply(
#             lambda row: determine_priority(row['Email'], row['firstName'], row['lastname'], row['webSites']),
#             axis=1
#         )
        
#         # Filter out rows without a valid priority
#         valid_rows = group[group['Priority'].notna()]
        
#         # If P1 exists, select it and discard others
#         p1_row = valid_rows[valid_rows['Priority'] == 'P1']
#         if not p1_row.empty:
#             processed_rows.append(p1_row.iloc[0])
#         else:
#             # If no P1, check for P2
#             p2_row = valid_rows[valid_rows['Priority'] == 'P2']
#             if not p2_row.empty:
#                 processed_rows.append(p2_row.iloc[0])

#     # Create a new DataFrame from the processed rows
#     processed_df = pd.DataFrame(processed_rows)
    
#     # Write the processed data to an output Excel file
#     processed_df.to_excel(output_path, index=False)
#     print(f"Processed data saved to: {output_path}")

# # Example usage
# input_file = r"C:\Users\Kompass\Downloads\Verified(750)_Email.xlsx"
# output_file = r"C:\Users\Kompass\Downloads\Priority_Verified(750)_Email.xlsx"
# process_excel(input_file, output_file)


























import pandas as pd

def determine_priority(email, first_name, last_name, domain):
    # Check if fields are valid strings
    if isinstance(email, str) and isinstance(first_name, str) and isinstance(last_name, str) and isinstance(domain, str):
        # Priority P1 for {first_name}.{last_name}@{domain}
        if email == f"{first_name.lower()}.{last_name.lower()}@{domain.lower()}":
            return 'P1'
        # Priority P2 for {first_name}@{domain}
        elif email == f"{first_name.lower()}@{domain.lower()}":
            return 'P2'
    # Eliminate other variations or invalid data
    return None

# Function to process the Excel data
def process_excel(file_path, output_path):
    # Read the Excel file into a DataFrame
    df = pd.read_excel(file_path)
    
    # Filter out inactive emails first (assuming 'active' column indicates this)
    active_df = df[df['Status'] == 'Active']  # Update 'active' column name as needed
    
    # Initialize a list to store the processed rows
    processed_rows = []
    
    # Group by 'webSites' to process each domain separately
    grouped = active_df.groupby(['webSites'])

    for _, group in grouped:
        # Iterate over each row in the group
        group['Priority'] = group.apply(
            lambda row: determine_priority(row['Email'], row['firstName'], row['lastName'], row['webSites']),
            axis=1
        )
        
        # Filter out rows without a valid priority
        valid_rows = group[group['Priority'].notna()]
        
        # If P1 exists, select it and discard others
        p1_row = valid_rows[valid_rows['Priority'] == 'P1']
        if not p1_row.empty:
            # Add extra fields to the row
            p1_row = p1_row[['_id', 'salutation', 'firstName', 'lastName', 'functionName', 'functionCode', 'linkedinprofile', 'Email', 'Status', 'Reason', 'Priority']]
            processed_rows.append(p1_row.iloc[0])
        else:
            # If no P1, check for P2
            p2_row = valid_rows[valid_rows['Priority'] == 'P2']
            if not p2_row.empty:
                # Add extra fields to the row
                p2_row = p2_row[['_id','salutation', 'firstName', 'lastName', 'functionName', 'functionCode', 'linkedinprofile', 'Email', 'Status', 'Reason', 'Priority']]
                processed_rows.append(p2_row.iloc[0])

    # Create a new DataFrame from the processed rows
    processed_df = pd.DataFrame(processed_rows)
    
    # Write the processed data to an output Excel file
    processed_df.to_excel(output_path, index=False)
    print(f"Processed data saved to: {output_path}")

# Example usage
input_file = r"C:\Users\Kompass\Desktop\Sanskriti Linkedin Email\Batches\Batch(1001-1403)Sanskruti_Linkedin_VerifiedEmail_Output_31.01.2025.xlsx.xlsx.xlsx"
output_file = r"C:\Users\Kompass\Desktop\Sanskriti Linkedin Email\Batches\Priority_Batch(1001-1403)Sanskruti_Linkedin_VerifiedEmail_Output_3.02.2025.xlsx.xlsx.xlsx"
process_excel(input_file, output_file)
