import pandas as pd
from pathlib import Path
import re
import PyPDF2
import os

DB_FILE = "campaign_database.xlsx"

def extract_data_from_pdf(file_path):
    def safe_extract(patterns, text):
        for pattern in patterns:
            match = re.search(pattern, text, re.MULTILINE)
            if match:
                # Check all possible capture groups
                for group in match.groups():
                    if group:
                        return int(group)
        return 0

    try:
        with open(file_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            if reader.is_encrypted:
                if not reader.decrypt(''):
                    print(f"Warning: Could not decrypt {file_path}. Skipping.")
                    return None

            text = ""
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text

            return {
                "Campaign": Path(file_path).stem,
                "Emails Sent": safe_extract([r"emails sent\s*(\d+)", r"(\d+)\s*\nTotal Emails Sent"], text),
                "Delivered": safe_extract([r"Delivered \d+\.\d+%\s*(\d+)", r"(\d+)\s*Delivered"], text),
                "Unique Opens": safe_extract([r"Unique Opens \d+\.\d+%\s*(\d+)", r"(\d+)\s*Opened \(Unique\)"], text),
                "Total Opens": safe_extract([r"Total Opens Count\s*(\d+)"], text),
                "Unique Clicks": safe_extract([r"Unique Clicks \d+\.\d+%\s*(\d+)", r"(\d+)\s*Clicked \(Unique\)"], text),
                "Unsubscribes": safe_extract([r"Unsubscribes \d+\.\d+%\s*(\d+)"], text),
                "Bounces": safe_extract([r"Bounces \d+\.\d+%\s*(\d+)", r"(\d+)\s*Bounced"], text),
                "Hard Bounces": safe_extract([r"Hard Bounce\s*(\d+)\s*Contacts"], text),
                "Soft Bounces": safe_extract([r"Soft Bounce\s*(\d+)\s*Contacts"], text),
                "Complaints": safe_extract([r"Complaints \d+\.\d+%\s*(\d+)"], text),
                "Forwards": safe_extract([r"Forwards\s*(\d+)"], text),
                "Mobile": safe_extract([r"Mobile (\d+)\s*%"], text),
                "Desktop": safe_extract([r"Computer (\d+)\s*%"], text),
                "Tablet": safe_extract([r"Tablet (\d+)\s*%"], text),
                "Locations": re.findall(r"(\d+) Opens from ([\w\s]+)", text)
            }
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        return None

def main():
    print("Starting data extraction from PDF files...")
    pdf_files = list(Path('.').glob('*.pdf'))
    
    if not pdf_files:
        print("No PDF files found in the current directory.")
        return

    all_data = [extract_data_from_pdf(f) for f in pdf_files]
    all_data = [data for data in all_data if data is not None]

    if not all_data:
        print("No data could be extracted from the PDFs.")
        return

    print(f"Successfully extracted data from {len(all_data)} PDF(s).")

    # Separate campaign data from location data
    campaign_df = pd.DataFrame(all_data).drop(columns=['Locations'])
    
    location_data = []
    for _, row in pd.DataFrame(all_data).iterrows():
        for count, country in row['Locations']:
            location_data.append({"Campaign": row['Campaign'], "Country": country.strip(), "Opens": int(count)})
    location_df = pd.DataFrame(location_data)

    # Save to Excel
    print(f"Saving data to {DB_FILE}...")
    with pd.ExcelWriter(DB_FILE) as writer:
        campaign_df.to_excel(writer, sheet_name="Campaign_Data", index=False)
        location_df.to_excel(writer, sheet_name="Location_Data", index=False)
    
    print(f"Database created successfully: {DB_FILE}")

if __name__ == '__main__':
    main()
