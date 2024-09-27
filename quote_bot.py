from google.cloud import vision_v1 as vision
from google.cloud import storage
import os
import re
import json
import openpyxl

# Set up your credentials
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = r"quotebot-436720-a7ba4b073b71.json"

# --- OCR Extraction Part ---
def async_detect_document(gcs_source_uri, gcs_destination_uri):
    """OCR with PDF as source file and output results to Google Cloud Storage."""
    client = vision.ImageAnnotatorClient()

    # Set up feature for document text detection (OCR)
    feature = vision.Feature(type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION)

    # Input file configuration (PDF file in GCS)
    gcs_source = vision.GcsSource(uri=gcs_source_uri)
    input_config = vision.InputConfig(gcs_source=gcs_source, mime_type='application/pdf')

    # Output file location in Google Cloud Storage
    gcs_destination = vision.GcsDestination(uri=gcs_destination_uri)
    output_config = vision.OutputConfig(gcs_destination=gcs_destination, batch_size=1)

    # Create the request for async file annotation
    async_request = vision.AsyncAnnotateFileRequest(features=[feature], input_config=input_config,
                                                    output_config=output_config)

    # Initiate the asynchronous request
    operation = client.async_batch_annotate_files(requests=[async_request])
    print(f'Processing file: {gcs_source_uri}')
    operation.result(timeout=10000)  # Waits until the request is complete

    print(f'OCR output stored at: {gcs_destination_uri}')


def process_pdfs_in_folders(bucket_name, input_folder, output_folder):
    """Process all PDFs in the input folder of a GCS bucket and store JSON results in the output folder."""
    storage_client = storage.Client()

    # Get the bucket
    bucket = storage_client.bucket(bucket_name)

    # List all PDF objects in the input folder
    blobs = bucket.list_blobs(prefix=input_folder)

    # Iterate through each PDF file in the input folder
    for blob in blobs:
        if blob.name.endswith('.pdf'):  # Ensure you're only processing PDF files
            gcs_source_uri = f'gs://{bucket_name}/{blob.name}'
            print(f"Found PDF: {gcs_source_uri}")

            # Generate the output path for the JSON file in the output folder
            output_blob_name = blob.name.replace(input_folder, output_folder).replace('.pdf', '.json')
            gcs_destination_uri = f'gs://{bucket_name}/{output_blob_name}'
            print(f"Storing OCR result at: {gcs_destination_uri}")

            # Perform OCR on the PDF and save results in the output folder
            async_detect_document(gcs_source_uri, gcs_destination_uri)


def download_ocr_results(bucket_name, output_folder):
    """Download and process all OCR results (JSON files) from the output folder."""
    storage_client = storage.Client()

    # Get the bucket
    bucket = storage_client.bucket(bucket_name)

    # List all JSON objects in the output folder
    blobs = bucket.list_blobs(prefix=output_folder)

    # Process each JSON file
    full_text = ''
    for blob in blobs:
        content = blob.download_as_string()

        # Debug: Check content before parsing
        print(f"Processing blob: {blob.name}")
        if not content or content.strip() == "":
            print(f"Warning: {blob.name} is empty or invalid.")
            continue  # Skip this blob if it's empty or invalid

        try:
            response = json.loads(content)
            # Check if 'responses' exists in the JSON response
            if 'responses' in response:
                for page in response['responses']:
                    # Check if 'fullTextAnnotation' exists in the response
                    if 'fullTextAnnotation' in page:
                        full_text += page['fullTextAnnotation'].get('text', '')
                    else:
                        print(f"No 'fullTextAnnotation' found in {blob.name}")
            else:
                print(f"No 'responses' found in {blob.name}")
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON from {blob.name}: {e}")
            continue  # Skip this blob if there's a decoding error

    return full_text


def extract_specific_data(text):
    """Extract specific data like name of insured, partners, qualified assistants, indemnity, excess, and profession from the OCR text."""

    # Regex patterns to capture the data
    name_pattern = r"Full title of Proposer[^:]*?\n([A-Za-z\s]+LLP)"  # Extracts the name of the insured
    qualified_assistants_pattern = r"Number of staff[^:]*?\n(\d+)\n"  # Pattern to extract the number of qualified assistants
    indemnity_pattern = r"Indemnity:\s?KSHs?,?\s?([\d,]+)"  # Matches Indemnity value
    excess_pattern = r"Excess:\s?KSHS\.\s?([\d,]+)"  # Matches Excess value
    profession_pattern = r"give a detailed description of the activities of the business to be covered\.\n([A-Z\s]+)\s"  # Extract Profession

    # Extracting the name of the insured
    name_match = re.search(name_pattern, text)
    insured_name = name_match.group(1) if name_match else "Not found"
    print(f"Name of Insured: {insured_name}")

    # Extracting the number of partners between the specific sections
    partner_section_pattern = r"PATRICK MMRIGI\nPosition(.+?)Qualifications"
    partner_section_match = re.search(partner_section_pattern, text, re.DOTALL)

    if partner_section_match:
        partner_section = partner_section_match.group(1)
        partner_count = len(re.findall(r"\bPARTNER\b", partner_section))
    else:
        partner_count = 0

    print(f"Number of Directors/Partners: {partner_count}")

    # Extracting the number of qualified assistants
    qualified_assistants_match = re.search(qualified_assistants_pattern, text)
    qualified_assistants = qualified_assistants_match.group(1) if qualified_assistants_match else "Not found"
    print(f"Qualified Assistants: {qualified_assistants}")

    # Extracting indemnity value
    indemnity_match = re.search(indemnity_pattern, text)
    indemnity = indemnity_match.group(1) if indemnity_match else "Not found"
    print(f"Indemnity: {indemnity}")

    # Extracting excess value
    excess_match = re.search(excess_pattern, text)
    excess = excess_match.group(1) if excess_match else "Not found"
    print(f"Excess: {excess}")

    # Extracting the profession
    profession_match = re.search(profession_pattern, text)
    profession = profession_match.group(1).strip() if profession_match else "Not found"
    print(f"Profession: {profession}")

    # Return extracted data as a dictionary
    return {
        "name_of_insured": insured_name,
        "partners": partner_count,
        "qualified_assistants": qualified_assistants,
        "indemnity": indemnity,
        "excess": excess,
        "profession": profession
    }

# --- Excel Population Part ---
def calculate_d17(c17):
    if c17 <= 1000000:
        return 1.050 / 100
    elif c17 <= 2000000:
        return 0.750 / 100
    elif c17 <= 5000000:
        return 0.450 / 100
    elif c17 <= 10000000:
        return 0.350 / 100
    elif c17 <= 20000000:
        return 0.225 / 100
    elif c17 <= 50000000:
        return 0.125 / 100
    else:
        return 0.125 / 100

def calculate_d20(c20):
    if c20 <= 1000000:
        return 100/100
    elif c20 <= 2500000:
        return 150/100
    elif c20 <= 5000000:
        return 190/100
    elif c20 <= 10000000:
        return 230/100
    elif c20 <= 20000000:
        return 275/100
    elif c20 <= 40000000:
        return 325/100
    elif c20 <=60000000:
        return 365/100
    elif c20 <= 80000000:
        return 400/100
    else:
        return 450/100

def calculate_d22(c22):
    professions = {
        "Opticians": 100,
        "Chemists": 100,
        "Accountants": 100,
        "Auditor": 100,
        "Attorneys": 100,
        "Architects": 135,
        "Civil Engineers": 135,
        "Quantity Surveyors": 135,
        "Dentists": 175,
        "Doctors": 175,
        "Surgeons": 175
    }
    return professions.get(c22, 1)  # Default to 1 if not found

def populate_excel_template(extracted_data, template_path, output_path):
    # Load the Excel template
    wb = openpyxl.load_workbook(template_path)
    sheet = wb.active

    # Populate the sheet based on extracted data
    sheet['B8'] = f"INSURED: {extracted_data['name_of_insured']}"
    sheet['C12'] = int(extracted_data['partners'] or 0)
    sheet['C13'] = int(extracted_data['qualified_assistants'] or 0)
    sheet['C14'] = 0  # Assuming no unqualified assistants in extracted data
    sheet['C15'] = 0  # Assuming no others in extracted data

    # Set static rates
    sheet['D12'] = 3000
    sheet['D13'] = 2500
    sheet['D14'] = 2000
    sheet['D15'] = 1000

    # Populate calculated fields in E column
    sheet['E12'] = int(sheet['C12'].value or 0) * int(sheet['D12'].value or 0)
    sheet['E13'] = int(sheet['C13'].value or 0) * int(sheet['D13'].value or 0)
    sheet['E14'] = int(sheet['C14'].value or 0) * int(sheet['D14'].value or 0)
    sheet['E15'] = int(sheet['C15'].value or 0) * int(sheet['D15'].value or 0)

    # Indemnity calculations
    c17 = int(extracted_data['indemnity'].replace(',', ''))
    d17 = calculate_d17(c17)
    sheet['C17'] = c17
    sheet['D17'] = d17
    sheet['E17'] = c17 * d17

    # Total calculations for E18
    sheet['E18'] = (
        int(sheet['E12'].value or 0) +
        int(sheet['E13'].value or 0) +
        int(sheet['E14'].value or 0) +
        int(sheet['E15'].value or 0) +
        int(sheet['E17'].value or 0)
    )

    # Indemnity-related values for row 20
    c20 = c17
    d20 = calculate_d20(c20)
    sheet['C20'] = c20
    sheet['D20'] = d20
    sheet['E20'] = d20 * sheet['E18'].value

    # Profession-based calculations
    c22 = extracted_data['profession']
    d22 = calculate_d22(c22)
    sheet['C22'] = c22
    sheet['D22'] = d22
    sheet['E22'] = d22 * sheet['E20'].value

    # Additional calculations for E24 - E32
    sheet['E24'] = sheet['E18'].value + sheet['E20'].value + sheet['E22'].value
    sheet['E26'] = sheet['E24'].value * 0.10
    sheet['E27'] = sheet['E24'].value * 0.10
    sheet['E28'] = sheet['E24'].value * 0.10
    sheet['E29'] = sheet['E24'].value + sheet['E26'].value + sheet['E27'].value + sheet['E28'].value
    sheet['E30'] = sheet['E29'].value * 0.45/100
    sheet['E31'] = 40
    sheet['E32'] = sheet['E29'].value + sheet['E30'].value + sheet['E31'].value

    # Save the populated workbook
    wb.save(output_path)
    print(f"Populated Excel saved to {output_path}")

    return output_path

def main():
    # Paths to template and output files
    template_path = r"C:\xampp\htdocs\KenyaRE Project\myprojectenv\documents\PROFESSIONAL INDEMNITY BLANK QUOTATION TEMPLATE.xlsx"
    output_excel_path = r"C:\xampp\htdocs\KenyaRE Project\myprojectenv\documents\Populated_Professional_Indemnity_Quotation.xlsx"

    # Bucket and folder details for OCR processing
    bucket_name = 'quote-bot-docs'
    input_folder = 'input_folder/'
    output_folder = 'output_folder/'

    # Step 1: Process PDFs in input folder and store JSON results in output folder
    process_pdfs_in_folders(bucket_name, input_folder, output_folder)

    # Step 2: Download and process OCR results from output folder
    extracted_text = download_ocr_results(bucket_name, output_folder)
    #print("Extracted Text:")  #printing out the text extracted from JSON files
    #print(extracted_text)

    # Step 3: Extract specific data from the OCR results
    extracted_data = extract_specific_data(extracted_text)

    # Step 4: Populate the Excel template with extracted data
    populated_excel_path = populate_excel_template(extracted_data, template_path, output_excel_path)

    # Step 5: Provide option to download as Excel
    user_choice = input("Download the output as Excel file? Enter 1 to confirm download ... ")

    if user_choice == '1':
        print(f"Populated Excel file downloaded to this path: {populated_excel_path}")
    else:
        print("Invalid choice. Please select 1 for Excel file Download.")

if __name__ == "__main__":
    main()
