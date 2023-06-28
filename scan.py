import os
import tkinter as tk
from googleapiclient.discovery import build
from google.oauth2 import service_account
import winsound

# Google Sheets API credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDS_FILE = 'scan-390611-25fc1daa0744.json'  # Path to your credentials JSON file
SHEET_ID = '1JvvmA_Lbr55XwYD2atYzeoAbD_ntOJxPx9IDKyCbK5c'  # ID of the Google Sheet

# Create GUI
root = tk.Tk()
root.title('Ni Fashion')

# Entry field
lbl_awb = tk.Label(root, text='AWB:')
lbl_awb.grid(row=0, column=0)
entry_awb = tk.Entry(root)
entry_awb.grid(row=0, column=1)

# Labels for displaying 3PL and company names
lbl_3pl = tk.Label(root, text='3PL:')
lbl_3pl.grid(row=1, column=0)
lbl_company = tk.Label(root, text='Company:')
lbl_company.grid(row=2, column=0)

# Label for displaying error or duplicate message
lbl_message = tk.Label(root, text='')
lbl_message.grid(row=3, columnspan=2)

# Google Sheets API authentication
creds = service_account.Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)

# Connect to Google Sheets API
service = build('sheets', 'v4', credentials=creds)

# Function to submit data
def submit_data(event=None):
    awb = entry_awb.get()

    # Get data from Data to check for duplicates
    data_range = 'Data!A:A'
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=data_range
    ).execute()
    data_values = result.get('values', [])

    # Check for duplicates in Data sheet
    duplicate = False
    for row in data_values:
        if len(row) > 0 and awb == row[0]:
            duplicate = True
            lbl_message.config(text='Duplicate AWB!')
            # Play duplicate validation sound
            winsound.PlaySound('sound3.wav', winsound.SND_FILENAME)
            break

    if not duplicate:
        # Get data from Manifest sheet for validation and corresponding names
        validation_range = 'Manifest!D:F'
        result = service.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=validation_range
        ).execute()
        validation_data = result.get('values', [])

        # Perform validation against Manifest sheet data and get corresponding names
        valid = False
        for row in validation_data:
            if len(row) > 2 and awb == row[0]:
                valid = True
                lbl_3pl.config(text='3PL: ' + row[2])  # Display 3PL name
                lbl_company.config(text='Company: ' + row[1])  # Display company name
                break

        if valid:
            # Append data to Data sheet
            values = [awb]
            body = {'values': [values]}
            service.spreadsheets().values().append(
                spreadsheetId=SHEET_ID,
                range='Data!A:A',
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            # Play true validation sound
            winsound.PlaySound('sound1.wav', winsound.SND_FILENAME)
            lbl_message.config(text='Data added successfully!')
        else:
            lbl_message.config(text='Invalid AWB!')
            # Play false validation sound
            winsound.PlaySound('sound2.wav', winsound.SND_FILENAME)

    # Clear entry field after submitting
    entry_awb.delete(0, tk.END)

# Submit button
btn_submit = tk.Button(root, text='Submit', command=submit_data)
btn_submit.grid(row=0, column=2)

# Bind Enter key to submit data
root.bind('<Return>', submit_data)

# Run the GUI
root.mainloop()
