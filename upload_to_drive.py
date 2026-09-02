import os
import glob
import base64
import requests
import json
import datetime
from dotenv import load_dotenv

load_dotenv()

APP_SCRIPT_URL = os.environ.get("APP_SCRIPT_URL")

def upload_file(filepath, custom_filename=None):
    if not APP_SCRIPT_URL:
        print("Error: APP_SCRIPT_URL not found in .env")
        return False
        
    filename = custom_filename or os.path.basename(filepath)
    print(f"Uploading {filename}...")
    
    with open(filepath, "rb") as f:
        file_data = f.read()
        
    b64_data = base64.b64encode(file_data).decode("utf-8")
    
    payload = {
        "fileData": b64_data,
        "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "fileName": filename
    }
    
    try:
        response = requests.post(APP_SCRIPT_URL, json=payload)
        response.raise_for_status()
        result = response.json()
        
        if result.get("status") == "success":
            print(f"Success! URL: {result.get('url')}")
            return True
        else:
            print(f"Failed to upload {filename}. Message: {result.get('message')}")
            return False
            
    except Exception as e:
        print(f"Exception during upload of {filename}: {e}")
        return False

def main():
    target_dir = os.path.join("GRAB", "hasil_custom")
    master_file = os.path.join(target_dir, "0MASTER.xlsx")
    
    if not os.path.exists(master_file):
        print(f"Error: Master file not found at {master_file}")
        return
        
    timestamp = datetime.datetime.now().strftime("%Y%m%d %H%M")
    upload_filename = f"Grab {timestamp}.xlsx"
    
    print(f"Found master file to upload: {upload_filename}")
    
    if upload_file(master_file, custom_filename=upload_filename):
        print(f"\nUpload complete: 1/1 files uploaded successfully.")
    else:
        print(f"\nUpload failed.")

if __name__ == "__main__":
    main()
