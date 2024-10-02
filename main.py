import os
import threading
from tkinter import Tk, Label, Button, Text, Listbox, Scrollbar, messagebox, filedialog, Toplevel
from tkinter import END, NORMAL
from dotenv import load_dotenv
import msal
import requests
import webbrowser

# Load environment variables from .env
load_dotenv()

# Azure AD App details
CLIENT_ID = os.getenv('CLIENT_ID')
REDIRECT_URI = os.getenv('REDIRECT_URI')

# The Microsoft Graph API endpoints for personal accounts
GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# MSAL PublicClientApplication for OAuth2
app = msal.PublicClientApplication(CLIENT_ID, authority="https://login.microsoftonline.com/common")

# Token variable
access_token = None
verification_uri = None  # Store verification URL
user_code = None  # Store user_code
selected_folder = None  # Store selected OneDrive folder

# Function to initiate the OAuth2 sign-in process
def sign_in():
    global verification_uri, user_code

    # Use MSAL to initiate device flow
    flow = app.initiate_device_flow(scopes=["Files.ReadWrite", "User.Read"])
    
    if "user_code" not in flow:
        messagebox.showerror("Error", "Failed to create device flow. Please try again.")
        return

    # Save the generated user_code and verification_uri for later use
    verification_uri = flow['verification_uri']
    user_code = flow['user_code']

    # Call function to open a new window and display the user_code and sign-in button
    open_sign_in_window(user_code, verification_uri)

    # Start a new thread to handle the token acquisition (to avoid blocking the GUI)
    threading.Thread(target=acquire_token_by_device_flow, args=(flow,)).start()

# Function to acquire the token in a separate thread
def acquire_token_by_device_flow(flow):
    global access_token

    result = app.acquire_token_by_device_flow(flow)

    # Check if the token was successfully acquired
    if "access_token" in result:
        access_token = result['access_token']
        token_text.delete(1.0, END)  # Clear previous token
        token_text.insert(END, access_token)  # Insert the new token
        messagebox.showinfo("Sign In", "Sign-in successful!")
        list_folders()  # Automatically fetch folders after login
    else:
        error_message = result.get('error_description', 'Unknown error occurred.')
        messagebox.showerror("Error", f"Failed to sign in: {error_message}")

# Function to open a pop-up window with the user_code and the sign-in link
def open_sign_in_window(user_code, verification_uri):
    # Create a new top-level window (pop-up)
    sign_in_window = Toplevel(root)
    sign_in_window.title("Sign In to Microsoft")

    # Display the user code in an interactive Text widget
    Label(sign_in_window, text="Enter this code on the Microsoft login page:").pack(pady=10)
    
    user_code_text = Text(sign_in_window, height=2, width=20, wrap="word")
    user_code_text.insert(END, user_code)  # Insert the user code
    user_code_text.config(state=NORMAL)  # Make it editable so the user can copy the text
    user_code_text.pack(pady=5)

    # Button to copy the code to clipboard
    Button(sign_in_window, text="Copy Code to Clipboard", command=lambda: copy_code_to_clipboard(user_code_text)).pack(pady=5)

    # Button to open the Microsoft login page in the default browser
    Button(sign_in_window, text="Go to Login Page", command=lambda: webbrowser.open(verification_uri)).pack(pady=5)

# Function to copy the user code to the clipboard
def copy_code_to_clipboard(text_widget):
    user_code = text_widget.get(1.0, END).strip()
    root.clipboard_clear()
    root.clipboard_append(user_code)
    messagebox.showinfo("Success", "User code copied to clipboard!")

# Function to list OneDrive folders
def list_folders():
    global access_token
    headers = {'Authorization': f'Bearer {access_token}'}
    response = requests.get(f"{GRAPH_ENDPOINT}/me/drive/root/children", headers=headers)

    if response.status_code == 200:
        folders = response.json().get('value', [])
        if not folders:
            messagebox.showinfo("Info", "No folders found in your OneDrive root directory.")
        else:
            folder_list.delete(0, 'end')  # Clear previous list
            for folder in folders:
                if folder.get('folder'):  # Ensure it's a folder
                    folder_list.insert('end', folder['name'])
    else:
        error_data = response.json()
        error_message = error_data.get('error', {}).get('message', 'Unknown error')
        messagebox.showerror("Error", f"Failed to retrieve folders: {response.status_code} - {error_message}")

# Function to select files and upload them to the selected OneDrive folder
def upload_files():
    global access_token, selected_folder
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
            messagebox.showinfo("Success", f"Uploaded {file_name} successfully.")
        else:
            print(f"Failed to upload {file_name}. Status code: {response.status_code}")
            messagebox.showerror("Error", f"Failed to upload {file_name}. Status code: {response.status_code}")

# Create the Tkinter GUI window
root = Tk()
root.title("OneDrive File Uploader")

# Sign-in button
Button(root, text="Sign In to OneDrive", command=sign_in).pack(pady=10)

# Label for token
Label(root, text="Access Token:").pack()

# Text widget to display the token
token_text = Text(root, height=4, width=80, wrap="word")
token_text.pack(pady=5)

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