import pandas as pd
import glob
import os

# Get a list of all Excel files in a directory, excluding temporary files starting with "~$"
excel_files = glob.glob(r"D:\VISHAL KOMPASS\web-scraping\Linkedin Company Prompt and Scrape Executives data\*.xlsx")
excel_files = [file for file in excel_files if not os.path.basename(file).startswith("~$")]

# Initialize an empty dataframe to hold the merged data
merged_data = pd.DataFrame()

# Iterate through each Excel file and merge into a single dataframe
for file in excel_files:
    if os.path.isfile(file):  # Check if the file exists
        try:
            df = pd.read_excel(file, engine='openpyxl')  # Specify the engine explicitly
            merged_data = pd.concat([merged_data, df], ignore_index=True)
        except Exception as e:
            print(f"Error reading {file}: {e}")
    else:
        print(f"File {file} does not exist or is not accessible.")

# Dropping duplicate rows
merged_data = merged_data.drop_duplicates()

# Write the merged data to a new Excel file
merged_data.to_excel(r"D:/VISHAL KOMPASS/web-scraping/Linkedin Company Prompt and Scrape Executives data/Sample_Data(10)LinkedInExecutives_28-03-2025.xlsx", index=False, engine='openpyxl')






# import pandas as pd
# from pathlib import Path
# import traceback

# # Define input and output paths
# input_dir = Path(r"D:\VISHAL KOMPASS\web-scraping\Extracted Files")
# output_file = Path(r"D:/VISHAL KOMPASS/web-scraping/Linkedin Company Prompt and Scrape Executives data/LinkedinExecutives_PerCompany_12-03-2025.xlsx")

# # Get a list of all Excel files in the directory, excluding temporary files
# excel_files = [file for file in input_dir.glob("*.xlsx") if not file.name.startswith("~$")]

# if not excel_files:
#     print("No valid Excel files found.")
# else:
#     print(f"Found {len(excel_files)} Excel files. Merging...")

#     # Read and merge files
#     dataframes = []
#     for file in excel_files:
#         try:
#             df = pd.read_excel(file, engine='openpyxl')
#             dataframes.append(df)
#             print(f"✔ Successfully read: {file.name}")
#         except Exception:
#             print(f"❌ Error reading {file.name}:\n{traceback.format_exc()}")

#     if dataframes:
#         merged_data = pd.concat(dataframes, ignore_index=True).drop_duplicates()

#         # Save the merged data
#         try:
#             merged_data.to_excel(output_file, index=False, engine='openpyxl')
#             print(f"✅ Merged data saved to {output_file}")
#         except Exception:
#             print(f"❌ Error saving the file:\n{traceback.format_exc()}")
#     else:
#         print("No valid data to merge.")












# import pandas as pd
# import glob
# import os
# import tkinter as tk
# from tkinter import filedialog

# def select_input_folder():
#     folder_path = filedialog.askdirectory()
#     input_folder_entry.delete(0, tk.END)
#     input_folder_entry.insert(tk.END, folder_path)

# def select_output_folder():
#     folder_path = filedialog.askdirectory()
#     output_folder_entry.delete(0, tk.END)
#     output_folder_entry.insert(tk.END, folder_path)

# def merge_excel_files():
#     input_folder = input_folder_entry.get()
#     output_folder = output_folder_entry.get()
#     output_filename = output_filename_entry.get()

#     excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
#     excel_files = [file for file in excel_files if not os.path.basename(file).startswith("~$")]

#     merged_data = pd.DataFrame()

#     for file in excel_files:
#         if os.path.isfile(file):
#             try:
#                 df = pd.read_excel(file, engine='openpyxl')
#                 merged_data = pd.concat([merged_data, df], ignore_index=True)
#             except Exception as e:
#                 print(f"Error reading {file}: {e}")
#         else:
#             print(f"File {file} does not exist or is not accessible.")

#     merged_data = merged_data.drop_duplicates()

#     output_file_path = os.path.join(output_folder, f"{output_filename}.xlsx")
#     merged_data.to_excel(output_file_path, index=False, engine='openpyxl')
#     message_text = f"Merged data saved to: {output_file_path}"
#     message_text_widget.insert(tk.END, message_text + "\n")

# # Create main window
# root = tk.Tk()
# root.title("Consolidation Manager")

# # Set window icon
# icon_path = r"C:/Users/Kompass/Downloads/download.jpg"  # Replace "path_to_your_icon.ico" with the path to your icon file

# from tkinter import PhotoImage

# try:
#   img = PhotoImage(file=icon_path)
#   root.iconphoto(False, img)  # Set icon for main window
# except tk.TclError as e:
#   print("Error setting icon:", e)

# # Input folder selection
# input_folder_label = tk.Label(root, text="Select Input Folder:")
# input_folder_label.grid(row=0, column=0, sticky=tk.W)

# input_folder_entry = tk.Entry(root, width=50)
# input_folder_entry.grid(row=0, column=1, padx=5, pady=5)

# input_folder_button = tk.Button(root, text="Browse", command=select_input_folder)
# input_folder_button.grid(row=0, column=2, padx=5, pady=5)

# # Output folder selection
# output_folder_label = tk.Label(root, text="Select Output Folder:")
# output_folder_label.grid(row=1, column=0, sticky=tk.W)

# output_folder_entry = tk.Entry(root, width=50)
# output_folder_entry.grid(row=1, column=1, padx=5, pady=5)

# output_folder_button = tk.Button(root, text="Browse", command=select_output_folder)
# output_folder_button.grid(row=1, column=2, padx=5, pady=5)

# # Output filename entry
# output_filename_label = tk.Label(root, text="Output Filename:")
# output_filename_label.grid(row=2, column=0, sticky=tk.W)

# output_filename_entry = tk.Entry(root, width=50)
# output_filename_entry.grid(row=2, column=1, padx=5, pady=5)

# # Message display area
# message_frame = tk.Frame(root)
# message_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=5)

# message_label = tk.Label(message_frame, text="Merge Status:")
# message_label.grid(row=0, column=0, sticky=tk.W)

# message_text_widget = tk.Text(message_frame, width=60, height=10)
# message_text_widget.grid(row=1, column=0, sticky=tk.W)

# scrollbar = tk.Scrollbar(message_frame, command=message_text_widget.yview)
# scrollbar.grid(row=1, column=1, sticky='nsew')
# message_text_widget.config(yscrollcommand=scrollbar.set)

# # Merge button
# merge_button = tk.Button(root, text="Merge Excel Files", command=merge_excel_files)
# merge_button.grid(row=4, column=1, pady=10)

# root.mainloop()
