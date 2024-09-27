# Quote-bot
Automation of Professional Indemnity Insurance Quotation with AI Bot


**Problem Statement**
The current process for handling Professional Indemnity Insurance quotation requests is manual, involving the labor-intensive task of analyzing proposal forms and other documents. This approach is time-consuming, tedious, and prone to delays, which can slow down the turnaround time for quotes and negatively impact business partner relationships. The reliance on manual analysis creates bottlenecks, particularly during periods of high demand, leading to inefficiencies and potential loss of business opportunities due to delayed responses. Automating this process is essential to improve efficiency, reduce errors, and ensure timely delivery of quotations to business partners.

**Proposed Solution Description (including tools used):**
The proposed solution utilizes a combination of Optical Character Recognition (OCR) and data automation tools to extract information from PDF documents and populate an Excel template with relevant data. Google Cloud Vision API is used for asynchronous document processing to extract specific information like the Name of Insured, Directors/Partners, Qualified Assistants, Indemnity, and Excess from PDFs. The extracted data is then processed using Python, with libraries like re for pattern matching and openpyxl for Excel manipulation. The program automatically fills in the appropriate cells in a blank Excel template, ensuring that data is entered based on predefined conditions and logic.

** The System Diagram (Workflow Description):**

**Input (PDF Upload):**
Users upload the PDF documents containing the required information.

**Google Cloud Vision API (OCR):**
The PDF is sent to Google Cloud Vision API, which extracts the text from the document asynchronously.
The API processes the document and returns the extracted data in JSON format.

**Data Extraction and Processing (Python):**
Using Python and regular expressions (re), the program processes the JSON response to extract specific values like Name of Insured, Partners, Qualified Assistants, Indemnity, and Excess.
Conditions are checked (e.g., number of partners).

**Excel Manipulation (openpyxl):**
The program loads the Excel template using openpyxl.
Based on the extracted values, specific cells in the template are populated.
The program then saves the updated Excel file with the populated data.

**Output (Excel File):**
The populated Excel file is saved and can be downloaded and printed for further use.

Expected Output
The expected output is a populated Excel file based on data extracted from PDF documents. The populated spreadsheet will contain key information such as the Name of Insured, the number of Directors/Partners, Qualified Assistants, Indemnity, and Excess. This automated workflow eliminates manual data entry, reduces errors, and provides a structured and consistent method of filling out professional indemnity quotation templates. The final Excel file is ready for review, if approved can be converted to pdf, downloaded, or printed, streamlining the insurance proposal process.


**Tools**
This is a comprehensive solution that integrates Google Cloud Vision API for OCR (optical character recognition), Google Cloud Storage for managing files, and Excel template population using OpenPyXL. Here's a breakdown of the functionalities of the tools:

**Key Functionalities:**
OCR with Google Cloud Vision API:
The code uploads PDF files stored in a Google Cloud Storage (GCS) bucket in the input_folder, processes them using Vision API for document text detection, and saves the results (as JSON) back to Google Cloud Service under the output_folder..

**Data Extraction:**
Specific data like name of insured, number of partners, indemnity value, etc., is extracted using regular expressions from the OCR text output.

**Excel Template Population:**
Extracted data is populated into an Excel template using OpenPyXL. Custom calculations for specific cells (indemnity, fees, etc.) are implemented through functions and pre defined expressions.
