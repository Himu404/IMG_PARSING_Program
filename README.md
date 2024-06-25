# OCR Image to Excel Converter

This Python program utilizes the Google Vision API to perform Optical Character Recognition (OCR) on images and parses the extracted data into an Excel file in a specific format.

## Features

- **Automated OCR**: Extracts text from images using Google Vision API.
- **Data Parsing**: Parses information like names, addresses, emails, and phone numbers.
- **Excel Output**: Outputs parsed data into a structured Excel workbook.
- **Customizable**: Easily configurable for different image directories and Google Cloud service keys.

## Setup Instructions

1. **Folder Path Configuration**:
   - Update `folder_path` variable in `main.py` to point to your image directory.
   
2. **Google Cloud Setup**:
   - Replace `json_key_path` with the path to your Google Cloud service account key JSON file.
   
3. **Dependencies**:
   - Install required Python libraries:
     ```bash
     pip install google-cloud-vision openpyxl
     ```

## Usage

1. Run the program:
   ```bash
   python main.py

    The program will process each image in the specified folder and generate parsed_information.xlsx with parsed data.

## Contributing

Contributions are welcome! Fork the repository, make your changes, and submit a pull request.

## Contact

For questions or suggestions, please contact najmulhaqe164@gmail.com.
