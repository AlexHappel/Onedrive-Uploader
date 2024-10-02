import os
from tkinter import Tk, ttk, filedialog, Label, Entry, Button, messagebox, Scrollbar, Listbox
from dotenv import load_dotenv
import msal
import requests
import webbrowser

# Load environment variables from .env
load_dotenv()

# Azure AD App details
CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
REDIRECT_URI = os.getenv('REDIRECT_URI')

# The authority URL for Microsoft (based on your Tenant ID)
AUTHORITY_URL = f'https://login.microsoftonline.com/{TENANT_ID}'

# The Microsoft Graph API endpoints
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# MSAL PublicClientApplication for OAuth2
app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY_URL)

# Token variable
access_token = None

# Function to initiate the OAuth2 sign-in process
def sign_in():
    global access_token
    flow = app.initiate_device_flow(scopes=["Files.ReadWrite.All", "User.Read"])

    if "user_code" not in flow:
        messagebox.showerror("Error", "Failed to create device flow. Please try again.")
        return

    # Show the user the code and provide the login URL
    messagebox.showinfo("Sign In", f"To sign in, go to {flow['verification_uri']} and enter the code: {flow['user_code']}")
    webbrowser.open(flow['verification_uri'])

    try:
        # Attempt to acquire token using the device flow
        result = app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            access_token = result['access_token']
            messagebox.showinfo("Sign In", "Sign-in successful!")
            list_folders()  # Automatically fetch folders after login
        else:
            # Log error details for debugging
            error_message = f"Error: {result.get('error')} - {result.get('error_description')}"
            print(error_message)  # Log to console
            messagebox.showerror("Error", error_message)
    except Exception as e:
        # Capture any unexpected exceptions and display them
        messagebox.showerror("Sign In Error", f"An unexpected error occurred: {str(e)}")

# Function to get a list of OneDrive folders
def list_folders():
    global access_token
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(f"{GRAPH_ENDPOINT}/me/drive/root/children", headers=headers)

    if response.status_code == 200:
        folders = response.json()['value']
        folder_list.delete(0, 'end')  # Clear previous list
        for folder in folders:
            if folder['folder']:  # Ensure it's a folder
                folder_list.insert('end', folder['name'])
    else:
        messagebox.showerror("Error", f"Failed to retrieve folders: {response.status_code}")

# Function to upload selected files to the selected OneDrive folder
def upload_files():
    global access_token
    selected_folder = folder_list.get(folder_list.curselection())
    
    if not selected_folder:
        messagebox.showerror("Error", "Please select a folder.")
        return

    files = filedialog.askopenfilenames()

    if not files:
        messagebox.showerror("Error", "Please select files to upload.")
        return

    headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/octet-stream'}

    for file_path in files:
        file_name = os.path.basename(file_path)
        upload_url = f"{GRAPH_ENDPOINT}/me/drive/root:/{selected_folder}/{file_name}:/content"
        with open(file_path, 'rb') as file_data:
            response = requests.put(upload_url, headers=headers, data=file_data)

        if response.status_code in [200, 201]:
            print(f"Uploaded {file_name} successfully.")
        else:
            print(f"Failed to upload {file_name}. Status code: {response.status_code}")

# Create the Tkinter GUI window
root = Tk()
root.title("OneDrive File Uploader")

# Sign-in button
Button(root, text="Sign In to OneDrive", command=sign_in).pack(pady=10)

# Label for OneDrive folder selection
Label(root, text="Select OneDrive Folder:").pack()

# Scrollable Listbox for displaying folders
scrollbar = Scrollbar(root)
scrollbar.pack(side="right", fill="y")

folder_list = Listbox(root, yscrollcommand=scrollbar.set, height=10)
folder_list.pack(pady=10)
scrollbar.config(command=folder_list.yview)

# Upload button
Button(root, text="Select Files and Upload", command=upload_files).pack(pady=20)

# Start the Tkinter event loop
root.mainloop()