import pandas as pd
import requests
import json
import tkinter as tk
from tkinter import filedialog

# Constants
ZOHO_API_URL = 'https://www.zohoapis.in/crm/v2/Leads/upsert'
ACCESS_TOKEN = '1000.6e6c903f21bac03809c4c49aebb19041.6db1d7cfb7f410e1bdf758324096906c'
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
    # Validate Mobile numbers (e.g., must be 10 digits)
    for index, row in df.iterrows():
        mobile = row.get('Mobile')
        if isinstance(mobile, str):
            if len(mobile) == 10 and mobile.isdigit():
                df.at[index, 'Mobile'] = mobile
            else:
                print(f"Invalid Mobile number at row {index}: {mobile} (Reason: not 10 digits or contains non-digit characters)")
                df.at[index, 'Mobile'] = None  # Set invalid data to None or handle as needed
        else:
            print(f"Invalid Mobile number at row {index}: {mobile} (Reason: not a string)")
            df.at[index, 'Mobile'] = None  # Set invalid data to None or handle as needed

    # Check for mandatory fields
    for index, row in df.iterrows():
        if pd.isna(row.get('Last_Name')):
            print(f"Missing Last_Name at row {index}")
            df.at[index, 'Last_Name'] = 'Unknown'  # Set a default value or handle as needed

    # Validate Owner field (must be a valid bigint)
    for index, row in df.iterrows():
        owner = row.get('Owner')
        try:
            owner_id = int(owner)
            df.at[index, 'Owner'] = owner_id
        except (ValueError, TypeError):
            print(f"Invalid Owner at row {index}: {owner}")
            df.at[index, 'Owner'] = None  # Set invalid data to None or handle as needed

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