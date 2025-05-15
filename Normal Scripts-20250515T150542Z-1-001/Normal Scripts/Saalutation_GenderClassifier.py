import pandas as pd
from transformers import pipeline

# Load the pre-trained gender classification model
gender_pipeline = pipeline("text-classification", model="padmajabfrl/Gender-Classification")

# Read the input Excel file
df = pd.read_excel(r"D:\VISHAL KOMPASS\web-scraping\Linkedin Company Prompt and Scrape Executives data\Format_Sample_Data(10)LinkedInExecutives_28-03-2025.xlsx")

# Ensure the column exists
if "firstName" not in df.columns:
    raise ValueError("Excel file must have a column named 'firstName'")

# Convert all values to strings & replace NaNs with empty strings
df["firstName"] = df["firstName"].fillna("").astype(str)

# Run inference
results = gender_pipeline(df["firstName"].tolist())

# Convert predictions: "Male" -> "Mr.", "Female" -> "Ms."
def format_gender(res):
    label = res["label"]
    return "Mr." if label == "Male" else "Ms." if label == "Female" else "Unknown"

df["salutation"] = [format_gender(res) for res in results]

# Save output to Excel
df.to_excel(r"D:\VISHAL KOMPASS\web-scraping\Linkedin Company Prompt and Scrape Executives data\Salutation_Sample_Data(10)LinkedInExecutives_28-03-2025.xlsx", index=False)

print("✅ Prediction completed! Results saved")
