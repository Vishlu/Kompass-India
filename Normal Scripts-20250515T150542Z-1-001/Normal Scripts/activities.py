"""
This is a Final Script for SEO Generation.

"""


import google.generativeai as genai
import pandas as pd
from tqdm import tqdm
from queue import Queue
import threading
import time

# Configure API Key
API_KEY = "AIzaSyAFZVvhp1DmElGer7WbwoYCehnFj8VRElQ"
genai.configure(api_key=API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')

def generate_seo_description(row):
    """Generate dynamic SEO description with fixed delay"""
    default_prompt = f"""
    Generate an SEO-optimized company description for [{row['Company Name']}]({row['Website']}).
    Write the description in 2 to 3 concise paragraphs.

    Paragraph 1: Introduce the company by stating its name, primary industry,
    location, and nature of business (e.g., manufacturer, exporter, distributor).

    Paragraph 2: List core products/services and highlight target markets/industries.
    Naturally integrate keywords like '[Location] [Product] [Business Type]'(e.g., 'Stainless Steel Valve Manufacturer in Germany').
    Use synonyms like 'supplier', 'vendor', or 'maker'. Include global certifications
    (e.g., ISO, CE) if available. Keep it concise and SEO-friendly.
    """

    try:
        response = model.generate_content(default_prompt)
        return response.text
    except Exception as e:
        return f"Error: {str(e)}"

def worker(q, results, progress, lock):
    """Thread worker function with rate control"""
    while not q.empty():
        try:
            idx, row = q.get_nowait()

            # Start critical section
            with lock:
                result = generate_seo_description(row)
                time.sleep(3)  # Fixed 3-second delay between requests
                results.append((idx, result))
                progress.update(1)

        except Exception as e:
            break
        finally:
            q.task_done()

def process_with_queue(df):
    """Queue-based processing with rate limiting"""
    q = Queue()
    for idx, row in df.iterrows():
        q.put((idx, row))

    results = []
    progress = tqdm(total=len(df), desc="Generating SEO Descriptions")
    lock = threading.Lock()  # For rate limiting

    # Start single thread (since we're rate limiting anyway)
    t = threading.Thread(target=worker, args=(q, results, progress, lock))
    t.start()

    # Wait for completion
    q.join()
    t.join()
    progress.close()

    # Reconstruct DataFrame
    for idx, result in results:
        df.at[idx, 'SEO_Description'] = result
    return df

def process_excel(input_file, output_file):
    """Main processing function"""
    df = pd.read_excel(input_file)

    # Validate required columns
    required_cols = ['Company Name', 'Website']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise ValueError(f"Missing columns: {', '.join(missing_cols)}")

    # Process with queue
    df = process_with_queue(df)

    # Save results
    df.to_excel(output_file, index=False)
    print(f"\n✅ Success! Processed {len(df)} companies")
    print(f"📁 Output saved to: {output_file}")

if __name__ == "__main__":
    process_excel(
        "/content/drive/MyDrive/KOMPASS INDIA UPLOAD/Excelsheet/Sample_Company_Description(activities)_29-03-2025.xlsx",
        "/content/drive/MyDrive/KOMPASS INDIA UPLOAD/Excelsheet/Output_Improved_Sample_Company_Description(activities)_1-04-2025.xlsx"
    )