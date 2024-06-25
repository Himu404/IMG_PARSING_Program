import io
import os
from google.cloud import vision
from openpyxl import Workbook

# Path to the folder containing image files
folder_path = r'C:\Users\Himu\Desktop\9627-9726'

# Path to your service account key JSON file
json_key_path = r'C:\Users\Himu\Desktop\New folder (3)\fleet-rite-427411-s1-a1c0b58f852c.json'

# Set the environment variable for authentication
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = json_key_path

# Instantiates a client
client = vision.ImageAnnotatorClient()

# Function to extract and parse information from an image
def extract_and_parse_image_info(image_path, ws, serial_number):
    image_name = os.path.basename(image_path)
    print(f"Processing image: {image_name} (Serial Number: {serial_number})")

    # Loads the image into memory
    with io.open(image_path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    # Performs text detection on the image file
    response = client.text_detection(image=image)
    texts = response.text_annotations

    # Check for errors in the response
    if response.error.message:
        print(f"Error in processing image {image_name}: {response.error.message}")
        return

    # Extract and parse only the first text annotation (formatted text)
    if texts:
        extracted_text = texts[0].description

        # Split the text into lines
        lines = extracted_text.split('\n')

        # Initialize variables to store the parsed data
        first_name = last_name = address = city = state = zip_code = email = owner_id = ""
        phones = {"Phone-1": "", "Phone-2": "", "Phone-3": ""}

        # Keywords for phone numbers (including variations)
        phone_keywords = ["phone1", "phone2", "phone 1", "phone 2", "phone 3", "home phone", "cell phone", "business"]

        # Iterate over lines to find relevant information
        for line in lines:
            # Convert line to lowercase for case-insensitive matching
            line_lower = line.lower()
            if line_lower.startswith("first name:"):
                first_name = line.split(":", 1)[1].strip()
            elif line_lower.startswith("last name:"):
                last_name = line.split(":", 1)[1].strip()
            elif line_lower.startswith("street address:"):
                address = line.split(":", 1)[1].strip()
            elif line_lower.startswith("city:"):
                city = line.split(":", 1)[1].strip()
            elif line_lower.startswith("state:"):
                state = line.split(":", 1)[1].strip()
            elif line_lower.startswith("zip:"):
                zip_code = line.split(":", 1)[1].strip()
            elif line_lower.startswith("email:"):
                email = line.split(":", 1)[1].strip()
            elif any(keyword in line_lower for keyword in phone_keywords):
                # Extract the phone number
                parts = line.split(":", 1)
                if len(parts) > 1:
                    phone_number = parts[1].strip()
                    # Check if the phone number already exists in phones
                    if phone_number not in phones.values():
                        # Find the first empty slot in phones to store the phone number
                        for key in phones:
                            if not phones[key]:
                                phones[key] = phone_number
                                break
            elif line_lower.startswith("owner id:"):
                owner_id = line.split(":", 1)[1].strip()

        # Print the parsed information to terminal
        print(f"\nParsed Information: ({image_name}) (Serial Number: {serial_number})")
        if first_name or last_name or address or city or state or zip_code or email or phones["Phone-1"] or phones["Phone-2"] or phones["Phone-3"] or owner_id:
            print(f"First Name: {first_name}")
            print(f"Last Name: {last_name}")
            print(f"Address: {address}")
            print(f"City: {city}")
            print(f"State: {state}")
            print(f"Zip: {zip_code}")
            print(f"E-mail: {email}")
            print(f"Phone-1: {phones['Phone-1']}")
            print(f"Phone-2: {phones['Phone-2']}")
            print(f"Phone-3: {phones['Phone-3']}")
            print(f"Owner ID: {owner_id}")

            # Write data to Excel worksheet
            ws.append([first_name, last_name, address, city, state, zip_code, email, phones["Phone-1"], phones["Phone-2"], phones["Phone-3"], owner_id, image_name])
        else:
            print("No relevant data found.")

# Create a new Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Parsed Information"

# Write header row
ws.append(["First Name", "Last Name", "Address", "City", "State", "Zip", "E-mail", "Phone-1", "Phone-2", "Phone-3", "Owner ID", "Image Name"])

# Iterate over all files in the folder
serial_number = 1
for filename in os.listdir(folder_path):
    # Check if the file is an image
    if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
        image_path = os.path.join(folder_path, filename)
        extract_and_parse_image_info(image_path, ws, serial_number)
        serial_number += 1

# Save the Excel workbook
wb.save("parsed_information.xlsx")
