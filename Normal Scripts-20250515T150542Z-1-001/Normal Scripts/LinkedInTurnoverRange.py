

# # Vishal code User Input Prompt for turnover range
# import re

# def classify_turnover(input_value):
#     # Define range categories in millions
#     categories = [
#         ("INR Under 5 Million", 0, 5),
#         ("INR 5 - 10 Million", 5, 10),
#         ("INR 10 - 50 Million", 10, 50),
#         ("INR 50 - 100 Million", 50, 100),
#         ("INR 100 - 500 Million", 100, 500),
#         ("INR 500 - 1000 Million", 500, 1000),
#         ("INR 1000 - 5000 Million", 1000, 5000),
#         ("INR 5000 - 10000 Million", 5000, 10000),
#         ("INR 10000 - 50000 Million", 10000, 50000),
#         ("Over 50000 Million", 50000, float('inf')),
#     ]

#     # Conversion factors
#     conversion_factors = {
#         "million": 1,
#         "cr": 10,  # 1 Cr = 10 Million
#     }

#     # Normalize the input
#     input_value = input_value.strip().lower()
#     input_value = re.sub(r"\s+", " ", input_value)  # Replace multiple spaces with a single space
#     input_value = input_value.replace("inr", "").strip()  # Remove INR part
#     input_value = input_value.replace(",", "")  # Remove commas
#     input_value = input_value.replace("–", "-")  # Normalize en-dash to hyphen "-"
    
#     # Handle "Under" and "Over"
#     multiplier = 1
#     if "under" in input_value:
#         multiplier = -1  # Indicates a maximum range
#         input_value = input_value.replace("under", "").strip()
#     elif "over" in input_value:
#         multiplier = 1  # Indicates a minimum range
#         input_value = input_value.replace("over", "").strip()

#     # Ensure that "cr" is properly separated from numbers
#     input_value = input_value.replace("cr", " cr")  # Adding a space before "cr" for consistency

#     # Handle ranges (e.g., "1 cr - 100 cr")
#     if "-" in input_value:
#         parts = input_value.split("-")
#         try:
#             # Process first and second part of the range
#             min_part = parts[0].strip()
#             max_part = parts[1].strip()

#             # Extract numeric value and units for both min and max
#             min_value = 0
#             max_value = 0

#             # Process first value (min_value)
#             for unit, factor in conversion_factors.items():
#                 if unit in min_part:
#                     numeric_part = min_part.split(unit)[0].strip()
#                     min_value = float(numeric_part.replace(" ", "")) * factor

#             # Process second value (max_value)
#             for unit, factor in conversion_factors.items():
#                 if unit in max_part:
#                     numeric_part = max_part.split(unit)[0].strip()
#                     max_value = float(numeric_part.replace(" ", "")) * factor

#             # Range Calculation: Average the two values
#             numeric_value = (min_value + max_value) / 2  # Take average for classification

#         except ValueError:
#             raise ValueError(f"Invalid range format: '{input_value}'")
#     else:
#         # Handle non-range cases
#         numeric_value = 0
#         for unit, factor in conversion_factors.items():
#             if unit in input_value:
#                 numeric_part = input_value.split(unit)[0].strip()
#                 try:
#                     numeric_value = float(numeric_part.replace(" ", "")) * factor
#                     break
#                 except ValueError:
#                     raise ValueError(f"Invalid numeric value: '{numeric_part}' in input '{input_value}'")

#     # Apply multiplier for "Under" or "Over"
#     if multiplier == -1:
#         numeric_value = max(0, numeric_value)  # For 'Under', we start from 0
#     elif multiplier == 1:
#         numeric_value += 0  # No adjustment needed for 'Over'

#     # Map numeric value to category
#     for category, lower, upper in categories:
#         if lower <= numeric_value < upper:
#             return category

#     return "Unknown category"

# # Examples of usage
# print(classify_turnover("8"))           # Should map to "INR Under 5 Million"
# print(classify_turnover("Over INR 10,000 Cr"))       # Should map to "Over 50000 Million"
# print(classify_turnover("Over INR 5,500 Cr"))        # Should map to "INR 10000 - 50000 Million"
# print(classify_turnover("Over INR 1,000 Cr"))        # Should map to "INR 10000 - 50000 Million"
# print(classify_turnover("80 Cr"))                     # Should map to "INR 50 - 100 Million"
# print(classify_turnover("INR 1 Cr – 100 Cr"))        # Should map to "INR 500 - 1000 Million"
# print(classify_turnover("INR 10 - 50 Million"))      # Should map to "INR 10 - 50 Million"










# # if column has blank cell it does not retun any value and final script
# import re
# import pandas as pd

# def classify_turnover(input_value):
#     # Check for blank or NaN values
#     if pd.isna(input_value) or input_value.strip() == "":
#         return ""  # Return blank for empty or NaN values

#     # Define range categories in millions
#     categories = [
#         ("INR Under 5 Million", 0, 5),
#         ("INR 5 - 10 Million", 5, 10),
#         ("INR 10 - 50 Million", 10, 50),
#         ("INR 50 - 100 Million", 50, 100),
#         ("INR 100 - 500 Million", 100, 500),
#         ("INR 500 - 1000 Million", 500, 1000),
#         ("INR 1000 - 5000 Million", 1000, 5000),
#         ("INR 5000 - 10000 Million", 5000, 10000),
#         ("INR 10000 - 50000 Million", 10000, 50000),
#         ("Over 50000 Million", 50000, float('inf')),
#     ]

#     # Conversion factors
#     conversion_factors = {
#         "million": 1,
#         "cr": 10,  # 1 Cr = 10 Million
#     }
 
#     # Normalize the input
#     input_value = input_value.strip().lower()
#     input_value = re.sub(r"\s+", " ", input_value)  # Replace multiple spaces with a single space
#     input_value = input_value.replace("inr", "").strip()  # Remove INR part
#     input_value = input_value.replace(",", "")  # Remove commas
#     input_value = input_value.replace("–", "-")  # Normalize en-dash to hyphen "-"

#     # Handle "Under" and "Over"
#     multiplier = 1
#     if "under" in input_value:
#         multiplier = -1  # Indicates a maximum range
#         input_value = input_value.replace("under", "").strip()
#     elif "over" in input_value:
#         multiplier = 1  # Indicates a minimum range
#         input_value = input_value.replace("over", "").strip()

#     # Ensure that "cr" is properly separated from numbers
#     input_value = input_value.replace("cr", " cr")  # Adding a space before "cr" for consistency

#     # Handle ranges (e.g., "1 cr - 100 cr")
#     if "-" in input_value:
#         parts = input_value.split("-")
#         try:
#             # Process first and second part of the range
#             min_part = parts[0].strip()
#             max_part = parts[1].strip()

#             # Extract numeric value and units for both min and max
#             min_value = 0
#             max_value = 0

#             # Process first value (min_value)
#             for unit, factor in conversion_factors.items():
#                 if unit in min_part:
#                     numeric_part = min_part.split(unit)[0].strip()
#                     min_value = float(numeric_part.replace(" ", "")) * factor

#             # Process second value (max_value)
#             for unit, factor in conversion_factors.items():
#                 if unit in max_part:
#                     numeric_part = max_part.split(unit)[0].strip()
#                     max_value = float(numeric_part.replace(" ", "")) * factor

#             # Range Calculation: Average the two values
#             numeric_value = (min_value + max_value) / 2  # Take average for classification

#         except ValueError:
#             raise ValueError(f"Invalid range format: '{input_value}'")
#     else:
#         # Handle non-range cases
#         numeric_value = 0
#         for unit, factor in conversion_factors.items():
#             if unit in input_value:
#                 numeric_part = input_value.split(unit)[0].strip()
#                 try:
#                     numeric_value = float(numeric_part.replace(" ", "")) * factor
#                     break
#                 except ValueError:
#                     raise ValueError(f"Invalid numeric value: '{numeric_part}' in input '{input_value}'")

#     # Apply multiplier for "Under" or "Over"
#     if multiplier == -1:
#         numeric_value = max(0, numeric_value)  # For 'Under', we start from 0
#     elif multiplier == 1:
#         numeric_value += 0  # No adjustment needed for 'Over'

#     # Map numeric value to category
#     for category, lower, upper in categories:
#         if lower <= numeric_value < upper:
#             return category

#     return "Unknown category"

# # Read Excel file
# excel_file_path = r"C:\Users\Kompass\Downloads\Final_DHL INDIA_Database Sourcing VIA Company List_DataSheets_KOMPASS_10.04.2025.xlsx"  # Replace with your file path
# df = pd.read_excel(excel_file_path)

# # Ensure the column with turnover data is named 'Turnover' (or change it as needed)
# df['Turnover KOM Range'] = df['turnover'].apply(classify_turnover)

# # Save the updated dataframe to a new Excel file
# output_file_path = r"C:\Users\Kompass\Downloads\Final2_DHL INDIA_Database Sourcing VIA Company List_DataSheets_KOMPASS_10.04.2025.xlsx"  # Set your desired output file name
# df.to_excel(output_file_path, index=False)

# print(f"Turnover classification completed and saved to {output_file_path}")














    