import pandas as pd
import pyqrcode
from PIL import ImageTk, Image, ImageDraw, ImageFont
import tkinter as tk
from tkinter import ttk

class QRSlideshowApp:
    def __init__(self, excel_file, sheet_name, column_name1, column_name2, qr_interval, combo_interval):
        self.excel_file = excel_file
        self.sheet_name = sheet_name
        self.column_name1 = column_name1
        self.column_name2 = column_name2
        self.qr_interval = qr_interval
        self.combo_interval = combo_interval
        self.qr_images = []
        self.current_index = 0

        self.window = tk.Tk()
        self.window.title("QR Code Slideshow")

        self.label_image = tk.Label(self.window)
        self.label_image.pack()

        self.label_text = tk.Label(self.window, font=("Arial", 12))
        self.label_text.pack()

        self.label_welcome = tk.Label(self.window, text="Ni Fashion Scanning System", font=("Arial", 16))
        self.label_welcome.pack(pady=10)

        self.progress_bar = ttk.Progressbar(self.window, length=200, mode='determinate')
        self.progress_bar.pack()

        self.start_button = tk.Button(self.window, text="Start", font=("Arial", 14), command=self.start_slideshow)
        self.start_button.pack(pady=10)

    def generate_qr_images(self):
        # Load data from Excel file
        df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name)

        # Extract data from specific columns in the DataFrame
        data_list1 = df[self.column_name1].tolist()
        data_list2 = df[self.column_name2].tolist()

        total_images = len(data_list1)
        loaded_images = 0

        for i in range(len(data_list1)):
            # Extract data for the pair
            data1 = str(data_list1[i])
            data2 = str(data_list2[i])

            # Generate QR codes
            qr_code1 = pyqrcode.create(data1)
            qr_code2 = pyqrcode.create(data2)

            # Save QR codes as images
            qr_code1.png(f"qrcode1.png", scale=6)
            qr_code2.png(f"qrcode2.png", scale=6)

            # Load images using PIL
            img1 = Image.open(f"qrcode1.png")
            img2 = Image.open(f"qrcode2.png")

            # Resize images (adjust the size as per your preference)
            img1 = img1.resize((250, 250))
            img2 = img2.resize((350, 350))

            # Add QR code value as text below the QR code image
            draw1 = ImageDraw.Draw(img1)
            draw2 = ImageDraw.Draw(img2)

            text1 = data1
            text2 = data2

            font = ImageFont.truetype("arial.ttf", 14)

            text_bbox1 = draw1.textbbox((0, 0), text1, font=font)
            text_width1 = text_bbox1[2] - text_bbox1[0]
            text_height1 = text_bbox1[3] - text_bbox1[1]

            text_bbox2 = draw2.textbbox((0, 0), text2, font=font)
            text_width2 = text_bbox2[2] - text_bbox2[0]
            text_height2 = text_bbox2[3] - text_bbox2[1]

            text_position1 = ((img1.width - text_width1) // 2, img1.height - text_height1 - 10)
            text_position2 = ((img2.width - text_width2) // 2, img2.height - text_height2 - 10)

            draw1.text(text_position1, text1, font=font, fill=0)
            draw2.text(text_position2, text2, font=font, fill=0)

            # Append QR code image pair to the list
            self.qr_images.append((ImageTk.PhotoImage(img1), ImageTk.PhotoImage(img2)))

            loaded_images += 1
            loading_percentage = int((loaded_images / total_images) * 100)
            self.progress_bar["value"] = loading_percentage
            self.window.update_idletasks()

        self.progress_bar.pack_forget()

    def update_images(self):
        qr_images_pair = self.qr_images[self.current_index]
        self.label_image.config(image=qr_images_pair[0])
        self.label_text.config(text="QR Code 1: " + self.get_data1() + "\nQR Code 2: " + self.get_data2())
        self.window.after(self.qr_interval * 1000, lambda: self.label_image.config(image=qr_images_pair[1]))
        self.window.after(self.qr_interval * 1000 + self.combo_interval * 1000, lambda: self.label_image.config(image=""))
        self.window.after(self.qr_interval * 1000 + self.combo_interval * 1000, self.next_image)

    def next_image(self):
        self.current_index = (self.current_index + 1) % len(self.qr_images)
        self.update_images()

    def get_data1(self):
        df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name)
        return str(df[self.column_name1].tolist()[self.current_index])

    def get_data2(self):
        df = pd.read_excel(self.excel_file, sheet_name=self.sheet_name)
        return str(df[self.column_name2].tolist()[self.current_index])

    def start_slideshow(self):
        self.start_button.config(state='disabled')  # Disable the start button
        self.generate_qr_images()
        self.update_images()
        self.window.mainloop()

# Example usage
excel_file = "data.xlsx"
sheet_name = "Sheet1"
column_name1 = "Data1"
column_name2 = "Data2"
qr_interval = 1  # Interval for displaying each QR code (in seconds)
combo_interval = 4  # Interval between each pair of QR codes (in seconds)

app = QRSlideshowApp(excel_file, sheet_name, column_name1, column_name2, qr_interval, combo_interval)
app.window.mainloop()
