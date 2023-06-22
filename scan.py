import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit
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
def manual_entry(app, main_window):
    barcode_input = QLineEdit()
    submit_button = QPushButton("Submit")
    quit_button = QPushButton("Quit")

    def submit_action():
        barcode_data = barcode_input.text()

        if barcode_data == 'q':
            app.quit()

        append_to_sheet(barcode_data)
        barcode_input.clear()

    submit_button.clicked.connect(submit_action)
    quit_button.clicked.connect(app.quit)

    widget = QWidget()
    layout = QVBoxLayout()
    layout.addWidget(QLabel("Enter Barcode (q to quit):"))
    layout.addWidget(barcode_input)
    layout.addWidget(submit_button)
    layout.addWidget(quit_button)
    widget.setLayout(layout)

    # Set the manual entry widget as the central widget of the main_window
    main_window.setCentralWidget(widget)

# Main function
def main():
    app = QApplication(sys.argv)
    main_window = QMainWindow()
    main_window.setWindowTitle("Barcode Scanner & Data Entry App")
    main_window.setGeometry(100, 100, 400, 300)

    scan_button = QPushButton("Scan Barcodes")
    manual_button = QPushButton("Manual Entry")
    quit_button = QPushButton("Quit")

    def scan_action():
        scan_barcodes()

    def manual_action():
        manual_entry(app, main_window)

    def quit_action():
        app.quit()

    scan_button.clicked.connect(scan_action)
    manual_button.clicked.connect(manual_action)
    quit_button.clicked.connect(quit_action)

    central_widget = QWidget(main_window)
    layout = QVBoxLayout()
    layout.addWidget(QLabel("Barcode Scanner & Data Entry App"))
    layout.addWidget(scan_button)
    layout.addWidget(manual_button)
    layout.addWidget(quit_button)
    central_widget.setLayout(layout)
    main_window.setCentralWidget(central_widget)
    main_window.show()

    sys.exit(app.exec())

# Run the app
if __name__ == '__main__':
    main()
