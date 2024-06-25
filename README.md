Certainly! Here's your README formatted in code blocks for easy copying into your GitHub repository:

```markdown
# OCR Image to Excel Converter

![Project Image](image_url)

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
     ```
     pip install google-cloud-vision openpyxl
     ```

## Usage

1. Clone the repository:
   ```
   git clone https://github.com/your_username/ocr-image-to-excel.git
   cd ocr-image-to-excel
   ```

2. Run the program:
   ```
   python main.py
   ```

3. The program will process each image in the specified folder and generate `parsed_information.xlsx` with parsed data.

## Example

Here's a sample command to run the program:
```
python main.py
```

## Contributing

Contributions are welcome! Fork the repository, make your changes, and submit a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contact

For questions or suggestions, please contact [Your Email Address].

---

Feel free to customize and expand upon this template based on your specific project details and preferences. Incorporate additional sections or modify existing ones as needed to provide comprehensive information about your OCR Image to Excel Converter project.
```

### Notes:
- Replace `image_url` with the URL of your project image or screenshot.
- Update paths and variable names (`folder_path`, `json_key_path`) according to your project's structure and requirements.
- Ensure that all commands and instructions are formatted correctly to maintain clarity and ease of use for your users.
