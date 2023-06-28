import cv2
from pyzbar import pyzbar
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import winsound
import tkinter as tk
from tkinter import filedialog

# Path to the Google Sheets credentials JSON file
CREDENTIALS_FILE = 'scan-390611-25fc1daa0744.json'

# Google Sheets document ID
DOCUMENT_ID = '1kTClIxeAcZhwXufiGUhxbJTn5X24fhIcjsxNS9Do9Ew'

# Connect to Google Sheets
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
client = gspread.authorize(credentials)
spreadsheet = client.open_by_key(DOCUMENT_ID)
sheet = spreadsheet.sheet1

# Create the Tkinter app
app = tk.Tk()
app.title("Barcode Scanner & Data Entry App")

# Function to append barcode data to Google Sheets and validate against sheet data
def append_to_sheet(barcode_data):
    # Open Sheet 2 within the Google Sheet document
    sheet_name = 'Sheet2'  # Replace with the name of Sheet 2
    worksheet = spreadsheet.worksheet(sheet_name)

    # Get all values from Sheet 2
    sheet_values = worksheet.get_all_values()

    # Search for the scanned barcode data in Sheet 2
    for row in sheet_values:
        if row[0] == barcode_data:
            # Print the corresponding 3PL name and company name
            if len(row) > 1:
                info_label.config(text="3PL Name: " + row[1])
            if len(row) > 2:
                company_name = row[2]
                info_label.config(text=info_label.cget("text") + "\nCompany Name: " + company_name)

                # Play sound effect based on company name if enable_company_sound is True
                if enable_company_sound:
                    if company_name == "Nayra":
                        winsound.PlaySound("Nayra.wav", winsound.SND_FILENAME)
                    elif company_name == "Ni":
                        winsound.PlaySound("Ni.wav", winsound.SND_FILENAME)
                    elif company_name == "DH":
                        winsound.PlaySound("DH.wav", winsound.SND_FILENAME)
                    # Add more conditions for other company names and corresponding sound effects

            # Check for duplicate entry in Sheet 1
            sheet1 = spreadsheet.sheet1
            existing_data = sheet1.get_all_values()
            if any(barcode_data in row for row in existing_data):
                info_label.config(text=info_label.cget("text") + "\nError: Duplicate entry.")
                winsound.PlaySound("error.wav", winsound.SND_FILENAME)
            else:
                # Append the barcode data to Sheet 1
                sheet1 = spreadsheet.sheet1
                row = [barcode_data]
                sheet1.append_row(row)
                info_label.config(text=info_label.cget("text") + "\nData saved successfully.")
            break
    else:
        # Scanned barcode data not found in Sheet 2 or did not validate
        info_label.config(text=info_label.cget("text") + "\nError: Barcode not found or data did not validate.")
        winsound.PlaySound("error.wav", winsound.SND_FILENAME)

# Barcode scanning function
def scan_barcodes():
    # Initialize camera
    cap = cv2.VideoCapture(0)

    while True:
        # Capture frame
        ret, frame = cap.read()

        # Decode barcodes
        barcodes = pyzbar.decode(frame)

        for barcode in barcodes:
            # Extract barcode data
            barcode_data = barcode.data.decode('utf-8')
            barcode_label.config(text="Scanned Barcode: " + barcode_data)
            append_to_sheet(barcode_data)

        # Display the frame
        cv2.imshow('Barcode Scanner', frame)

        # Wait for 'q' key to exit
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Release the capture and close the window
    cap.release()
    cv2.destroyAllWindows()

# Manual barcode entry function
def manual_entry(event=None):
    barcode_data = barcode_entry.get()
    if barcode_data == '':
        info_label.config(text="Error: Empty barcode.")
        winsound.PlaySound("error.wav", winsound.SND_FILENAME)
    else:
        append_to_sheet(barcode_data)
        barcode_entry.delete(0, tk.END)

# Toggle company sound function
def toggle_company_sound():
    global enable_company_sound
    enable_company_sound = not enable_company_sound
    if enable_company_sound:
        info_label.config(text=info_label.cget("text") + "\nCompany sound enabled.")
    else:
        info_label.config(text=info_label.cget("text") + "\nCompany sound disabled.")

# Create the GUI elements
barcode_label = tk.Label(app, text="Scanned Barcode: ")
barcode_label.pack()

info_label = tk.Label(app, text="")
info_label.pack()

scan_button = tk.Button(app, text="Scan Barcodes", command=scan_barcodes)
scan_button.pack()

manual_entry_frame = tk.Frame(app)
manual_entry_frame.pack()

manual_entry_label = tk.Label(manual_entry_frame, text="Manual Entry: ")
manual_entry_label.pack(side=tk.LEFT)

barcode_entry = tk.Entry(manual_entry_frame)
barcode_entry.pack(side=tk.LEFT)
barcode_entry.bind('<Return>', manual_entry)  # Bind Enter key press to manual_entry function

toggle_sound_button = tk.Button(app, text="Toggle Company Sound", command=toggle_company_sound)
toggle_sound_button.pack()

# Initialize the company sound status
enable_company_sound = True

# Start the Tkinter event loop
app.mainloop()
