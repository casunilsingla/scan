import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.google.zxing.BinaryBitmap;
import com.google.zxing.DecodeHintType;
import com.google.zxing.MultiFormatReader;
import com.google.zxing.NotFoundException;
import com.google.zxing.Reader;
import com.google.zxing.Result;
import com.google.zxing.common.HybridBinarizer;
import com.google.zxing.client.j2se.BufferedImageLuminanceSource;
import com.google.zxing.client.j2se.ImageReader;

import javax.sound.sampled.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.Collections;

public class BarcodeScannerApp {

    // Path to the Google Sheets credentials JSON file
    private static final String CREDENTIALS_FILE = "scan-390611-25fc1daa0744.json";

    // Google Sheets document ID
    private static final String DOCUMENT_ID = "1kTClIxeAcZhwXufiGUhxbJTn5X24fhIcjsxNS9Do9Ew";

    // Function to append barcode data to Google Sheets and validate against sheet data
    private static void appendToSheet(String barcodeData) throws IOException {
        // Create a Google credential from the service account credentials JSON file
        GoogleCredential credential = GoogleCredential.fromStream(new FileInputStream(CREDENTIALS_FILE))
                .createScoped(Collections.singleton(SheetsScopes.SPREADSHEETS));

        // Create a Google Sheets service
        Sheets sheetsService = new Sheets.Builder(credential.getTransport(), credential.getJsonFactory(), credential)
                .setApplicationName("Barcode Scanner App")
                .build();

        // Open Sheet 2 within the Google Sheet document
        String sheetName = "Sheet2"; // Replace with the name of Sheet 2

        // Get all values from Sheet 2
        ValueRange response = sheetsService.spreadsheets().values()
                .get(DOCUMENT_ID, sheetName + "!A1:C")
                .execute();

        List<List<Object>> sheetValues = response.getValues();

        // Search for the scanned barcode data in Sheet 2
        for (List<Object> row : sheetValues) {
            if (row.get(0).toString().equals(barcodeData)) {
                // Print the corresponding 3PL name and company name
                if (row.size() > 1) {
                    System.out.println("3PL Name: " + row.get(1));
                }
                if (row.size() > 2) {
                    String companyName = row.get(2).toString();
                    System.out.println("Company Name: " + companyName);

                    // Play sound effect based on company name
                    playSoundEffect(companyName);
                }

                // Check for duplicate entry in Sheet 1
                ValueRange existingData = sheetsService.spreadsheets().values()
                        .get(DOCUMENT_ID, "Sheet1!A:A")
                        .execute();

                List<List<Object>> existingValues = existingData.getValues();

                boolean isDuplicate = false;
                for (List<Object> rowData : existingValues) {
                    if (rowData.get(0).toString().equals(barcodeData)) {
                        isDuplicate = true;
                        break;
                    }
                }

                if (isDuplicate) {
                    System.out.println("Error: Duplicate entry.");
                    // Play sound effect for duplicate entry
                    playSoundEffect("error");
                } else {
                    // Append the barcode data to Sheet 1
                    sheetsService.spreadsheets().values()
                            .append(DOCUMENT_ID, "Sheet1", new ValueRange().setValues(Collections.singletonList(Collections.singletonList(barcodeData))))
                            .setValueInputOption("RAW")
                            .execute();

                    System.out.println("Data saved successfully.");
                }
                break;
            }
        }
        // Scanned barcode data not found in Sheet 2 or did not validate
        System.out.println("Error: Barcode not found or data did not validate.");
        // Play sound effect for invalid data
        playSoundEffect("error");
    }

    // Barcode scanning function
    private static void scanBarcodes() throws IOException, NotFoundException {
        // Initialize the camera or capture device

        // Initialize the barcode reader
        Reader reader = new MultiFormatReader();

        while (true) {
            // Capture frame or image

            // Decode barcodes
            BufferedImage bufferedImage = null; // Replace with the captured frame or image

            BinaryBitmap binaryBitmap = new BinaryBitmap(new HybridBinarizer(new BufferedImageLuminanceSource(bufferedImage)));
            Result[] barcodes = reader.decodeMultiple(binaryBitmap);

            for (Result barcode : barcodes) {
                // Extract barcode data
                String barcodeData = barcode.getText();
                System.out.println("Scanned Barcode: " + barcodeData);
                appendToSheet(barcodeData);
            }

            // Display the frame or image

            // Wait for 'q' key to exit
            // if (key == 'q') break;
        }

        // Release the capture device or camera
        // Close the window or UI component
    }

    // Manual barcode entry function
    private static void manualEntry() throws IOException {
        while (true) {
            // Read barcode data from user input

            // if (barcodeData.equals("q")) break;

            appendToSheet(barcodeData);
        }
    }

    // Play sound effect based on company name
    private static void playSoundEffect(String companyName) {
        try {
            String soundFileName;
            if (companyName.equals("Nayra")) {
                soundFileName = "Nayra.wav";
            } else if (companyName.equals("Ni")) {
                soundFileName = "Ni.wav";
            } else if (companyName.equals("DH")) {
                soundFileName = "DH.wav";
            } else {
                soundFileName = "error.wav";
            }

            File soundFile = new File(soundFileName);
            AudioInputStream audioInputStream = AudioSystem.getAudioInputStream(soundFile);
            Clip clip = AudioSystem.getClip();
            clip.open(audioInputStream);
            clip.start();
        } catch (UnsupportedAudioFileException | IOException | LineUnavailableException e) {
            e.printStackTrace();
        }
    }

    // Main function
    public static void main(String[] args) {
        System.out.println("Barcode Scanner & Data Entry App");
        System.out.println("1. Scan Barcodes");
        System.out.println("2. Manual Entry");

        // Read user's choice

        if (choice.equals("1")) {
            try {
                scanBarcodes();
            } catch (IOException | NotFoundException e) {
                e.printStackTrace();
            }
        } else if (choice.equals("2")) {
            try {
                manualEntry();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("Invalid choice. Exiting...");
        }
    }
}
