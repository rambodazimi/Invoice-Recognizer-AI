# Invoice-Recognizer-AI
A machine learning model to extract data from any invoice using Azure AI

## Invoice Data Extraction Script
This script utilizes Azure Form Recognizer to extract data from invoices in PDF format. It retrieves specific fields and their corresponding values from the invoices and saves the extracted data into a CSV file for further analysis or processing.

## Prerequisites
Before running the script, make sure you have the following:

Azure Form Recognizer endpoint and key: You need to create an Azure Form Recognizer resource in the Azure portal and obtain the endpoint URL and API key.
Installation
To use the script, follow these steps:

Clone the repository or download the script file to your local machine.

Install the required Python packages. You can use pip to install them:

```
pip install azure-ai-formrecognizer
```
Usage
Set the endpoint and key variables in the script with the values obtained from the Azure portal:

```
endpoint = "https://your-form-recognizer-endpoint-url/"
key = "your-form-recognizer-api-key"
```
Specify the URL of the invoice you want to analyze:

```
invoiceUrl = "https://path-to-your-invoice.pdf"
```
Run the script:

```
python invoice_extraction_script.py
```
The script will connect to the Azure Form Recognizer service, analyze the invoice, and extract the desired fields. The extracted data will be saved to a CSV file named extracted_data.csv.

## Output
The extracted data will be organized in a CSV file with the following structure:

Each invoice is separated by a row of dashes (--------Recognizing invoice #X--------).

The field values and their corresponding confidence scores are listed below the invoice header row.

The invoice items, such as line items or product details, are listed below the field values. Each item is separated by a row with the label Item #X, followed by the specific item details.

The invoices are separated by an empty row, followed by a row of dashes (----------------------------------------).

## Example
For a better understanding, here's an example of how the extracted data might appear in the CSV file:

```
--------Recognizing invoice #1--------
Vendor Name, Example Vendor
Vendor Address, 123 Main Street, City, State, ZIP
...
Invoice Id, INV-001
Invoice Date, 2023-05-15
Invoice Total, $500.00
...
Invoice items:
Item #1
Description, Product A
Quantity, 2
Unit Price, $100.00
...
Item #2
Description, Product B
Quantity, 1
Unit Price, $150.00
...
```
----------------------------------------
## Notes
The script supports analyzing multiple invoices in sequence. Each invoice will be separated in the CSV file for clarity.

If you run the script multiple times, the new data will be appended to the existing CSV file, ensuring a cumulative record of all processed invoices.

Make sure you have write permissions for the directory where the script is located to allow saving the CSV file.

## Acknowledgments
This script utilizes the Azure Form Recognizer service provided by Microsoft Azure. Learn more about Azure Form Recognizer and its capabilities at Azure Form Recognizer Documentation.

## Contact
If you have any questions or suggestions regarding this script, feel free to contact me at rambod.azm@gmail.com or visit my personal website at rambodazimi.com
