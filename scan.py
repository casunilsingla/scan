import cv2
from pyzbar import pyzbar
import gspread
from oauth2client.service_account import ServiceAccountCredentials

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

# Function to append barcode data to Google Sheets
def append_to_sheet(barcode_data):
    row = [barcode_data]
    sheet.append_row(row)
    print("Data saved successfully.")

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
