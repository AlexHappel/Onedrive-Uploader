import os
from tkinter import Tk, filedialog, Label, Entry, Button, messagebox
from dotenv import load_dotenv
import msal
import requests

# Load environment variables from .env
load_dotenv()

# Azure AD App details (now stored in env variables)
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')

# The authority URL for Microsoft (based on your Tenant ID)
AUTHORITY_URL = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPES = ['https://graph.microsoft.com/.default']

# MSAL ConfidentialClientApplication for OAuth2
app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY_URL, client_credential=CLIENT_SECRET
)

# Function to acquire an access token
def get_access_token():
    result = app.acquire_token_for_client(scopes=SCOPES)
    if 'access_token' in result:
        return result['access_token']
    else:
        print("Failed to acquire token:", result.get("error"), result.get("error_description"))
        return None

# Function to upload a file to OneDrive
def upload_file_to_onedrive(file_path, folder_path, token):
    file_name = os.path.basename(file_path)
    one_drive_upload_url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}/{file_name}:/content"

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/octet-stream'
    }

    with open(file_path, 'rb') as file_data:
        response = requests.put(one_drive_upload_url, headers=headers, data=file_data)

    if response.status_code in [201, 200]:
        return f"File {file_name} uploaded successfully!"
    else:
        return f"Failed to upload {file_name}. Status code: {response.status_code}, Response: {response.text}"

# Function to handle the file upload process
def start_upload():
    # Get the OneDrive folder path from the input field
    folder_path = folder_entry.get()

    # Get the selected files from the file dialog
    files = filedialog.askopenfilenames()

    if not folder_path or not files:
        messagebox.showerror("Error", "Please select files and specify the OneDrive folder.")
        return

    # Acquire access token
    token = get_access_token()
    if not token:
        messagebox.showerror("Error", "Failed to authenticate. Please check your credentials.")
        return

    # Upload each file
    for file_path in files:
        result = upload_file_to_onedrive(file_path, folder_path, token)
        print(result)
        messagebox.showinfo("Upload Status", result)

# Create the Tkinter GUI window
root = Tk()
root.title("OneDrive File Uploader")

# UI components
Label(root, text="OneDrive Folder Path:").pack(pady=10)
folder_entry = Entry(root, width=50)
folder_entry.pack(pady=5)

Button(root, text="Select Files and Upload", command=start_upload).pack(pady=20)

# Start the Tkinter event loop
root.mainloop()



