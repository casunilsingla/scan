import cv2
from pyzbar import pyzbar
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import winsound

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

import winsound

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
                print("3PL Name:", row[1])
            if len(row) > 2:
                company_name = row[2]
                print("Company Name:", company_name)

                # Play sound effect based on company name
                if company_name == "Nayra":
                    winsound.PlaySound("Nayra.wav", winsound.SND_FILENAME)  # Replace with the path to your sound effect file for Company A in WAV format
                elif company_name == "Ni":
                    winsound.PlaySound("Ni.wav", winsound.SND_FILENAME)  # Replace with the path to your sound effect file for Company B in WAV format
                elif company_name == "DH":
                    winsound.PlaySound("DH.wav", winsound.SND_FILENAME)  # Replace with the path to your sound effect file for Company B in WAV format
                # Add more conditions for other company names and corresponding sound effects

            # Check for duplicate entry in Sheet 1
            sheet1 = spreadsheet.sheet1
            existing_data = sheet1.get_all_values()
            if any(barcode_data in row for row in existing_data):
                print("Error: Duplicate entry.")
                # Play sound effect for duplicate entry
                winsound.PlaySound("error.wav", winsound.SND_FILENAME)  # Replace with the path to your sound effect file for duplicate entry in WAV format
            else:
                # Append the barcode data to Sheet 1
                sheet1 = spreadsheet.sheet1
                row = [barcode_data]
                sheet1.append_row(row)
                print("Data saved successfully.")
            break
    else:
        # Scanned barcode data not found in Sheet 2 or did not validate
        print("Error: Barcode not found or data did not validate.")
        # Play sound effect for invalid data
        winsound.PlaySound("error.wav", winsound.SND_FILENAME)  # Replace with the path to your sound effect file for invalid data in WAV format

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
            print("Scanned Barcode:", barcode_data)
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
def manual_entry():
    while True:
        barcode_data = input("Enter Barcode (q to quit): ")

        if barcode_data == 'q':
            break

        append_to_sheet(barcode_data)

# Main function
def main():
    print("Barcode Scanner & Data Entry App")
    print("1. Scan Barcodes")
    print("2. Manual Entry")

    choice = input("Select an option: ")

    if choice == '1':
        scan_barcodes()
    elif choice == '2':
        manual_entry()
    else:
        print("Invalid choice. Exiting...")

# Run the app
main()
