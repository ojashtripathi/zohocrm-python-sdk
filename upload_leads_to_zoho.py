import pandas as pd
import requests
import json
import tkinter as tk
from tkinter import filedialog

# Constants
ZOHO_API_URL = 'https://www.zohoapis.in/crm/v2/Leads/upsert'
ACCESS_TOKEN = '1000.7cb0360069d8e6e4066f0c6b517a9131.5a983964f8a8e6ce6bf5f9061083d7f4'
BATCH_SIZE = 100

# Function to upload a batch of leads to Zoho CRM
def upload_to_zoho(batch):
    headers = {
        'Authorization': f'Zoho-oauthtoken {ACCESS_TOKEN}',
        'Content-Type': 'application/json'
    }
    data = {
        'data': batch
    }
    response = requests.post(ZOHO_API_URL, headers=headers, data=json.dumps(data))
    if response.status_code == 200:
        print(f'Successfully uploaded batch of {len(batch)} records.')
    else:
        print(f'Failed to upload batch. Status code: {response.status_code}')
        print(response.text)

# Function to select an Excel file
def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

# Function to validate data
def validate_data(df):
    # Example: Validate Mobile numbers (this should be according to your specific requirements)
    for index, row in df.iterrows():
        if not isinstance(row['Mobile'], str) or not row['Mobile'].isdigit():
            print(f"Invalid Mobile number at row {index}: {row['Mobile']}")
            df.at[index, 'Mobile'] = None  # Set invalid data to None or handle as needed
    return df

def main():
    # Ask for the Excel file
    print("Please select the Excel file containing the leads.")
    excel_file_path = select_file()
    if not excel_file_path:
        print("No file selected. Exiting.")
        return

    # Read the Excel file
    df = pd.read_excel(excel_file_path)

    # Validate data
    df = validate_data(df)

    # Split the data into batches and upload each batch
    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i:i+BATCH_SIZE].to_dict(orient='records')
        upload_to_zoho(batch)

if __name__ == "__main__":
    main()