import tkinter as tk
from tkinter import filedialog
import requests
import csv
import os
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient
import threading
import subprocess
import pyautogui


# Set the endpoint and key variables with the values from the Azure portal
endpoint = "https://walterforms.cognitiveservices.azure.com/"
key = "d436aeca36244041b2c3cf21bee5921b"

# global variables
csv_filep = "C:\\Users\\razimi\\OneDrive - Walter Surface Technologies\\Desktop\\python\\form_recognizer\\output_prebuilt.csv"

def analyze_url():
    url = url_entry.get()

    # Start a new thread to perform the analysis
    analysis_thread = threading.Thread(target=perform_analysis, args=(url,))
    analysis_thread.start()
    
    try:
        message_label.config(text="PDF file in the url has not been found!", fg="red")
        response = requests.get(url)
        if response.status_code == 200:
            analyze_invoice(url, True, "noFile")
            message_label.config(text="CSV file has been generated successfully!!", fg="green")
        else:
            message_label.config(text="URL is invalid or inaccessible.", fg="red")

    except requests.exceptions.RequestException as e:
        message_label.config(text="Invalid URL!", fg="red")

def browse_pdf():
    message_label.config(text="Analyzing the document...", fg="blue")
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    
    if file_path:
        # Start a new thread to perform the analysis
        analysis_thread = threading.Thread(target=perform_analysis, args=(file_path,))
        analysis_thread.start()
        analyze_invoice("noURL", False, file_path)
        message_label.config(text="CSV file has been generated successfully!", fg="green")
        open_files_side_by_side(pdf_path=file_path, excel_path=csv_filep)


def perform_analysis(input_data):
    # Simulate a long-running analysis task
    import time
    time.sleep(5)  # Replace this with your actual analysis code
    

# Create the main window
window = tk.Tk()
window.title("Form Recognizer")
window.configure(bg="#F0F0F0")  # Set the background color of the window

# Create a logo label at the top
logo = tk.PhotoImage(file="C:\\Users\\razimi\\OneDrive - Walter Surface Technologies\\Desktop\\python\\form_recognizer\\icon.png")  # Replace "logo.png" with your actual logo file
logo_label = tk.Label(window, image=logo, bg="#F0F0F0")  # Set the background color of the logo label
logo_label.pack()

# Create a frame for the content
content_frame = tk.Frame(window, bg="#F0F0F0")  # Set the background color of the frame
content_frame.pack(pady=10)

# Create a heading label
heading_label = tk.Label(content_frame, text="Form Recognizer", font=("Arial", 16, "bold"), bg="#F0F0F0")  # Set the background color of the heading label
heading_label.pack()

# Create a label and entry for the URL
url_label = tk.Label(content_frame, text="Enter URL:", font=("Arial", 12), bg="#F0F0F0")  # Set the background color of the URL label
url_label.pack(pady=5)
url_entry = tk.Entry(content_frame, font=("Arial", 12))
url_entry.pack(pady=5)

# Create a button to trigger the analysis
analyze_button = tk.Button(content_frame, text="Analyze", font=("Arial", 12), command=analyze_url)
analyze_button.pack(pady=15)

# Create a button to browse for a PDF file
browse_button = tk.Button(content_frame, text="Browse", font=("Arial", 12), command=browse_pdf)
browse_button.pack(pady=10)

# Create a button to browse for a PDF file
# start_button = tk.Button(content_frame, text="Start", font=("Arial", 12))
# start_button.pack(pady=5)

# Create a message label
message_label = tk.Label(window, text="", font=("Arial", 12), bg="#F0F0F0")  # Set the background color of the message label
message_label.pack()

# Create a footer label
footer_label = tk.Label(window, text="© 2023 Walter Surface Technologies. All rights reserved.", font=("Arial", 10), bg="#F0F0F0")  # Set the background color of the footer label
footer_label.pack()

def format_bounding_region(bounding_regions):
    if not bounding_regions:
        return "N/A"
    return ", ".join("Page #{}: {}".format(region.page_number, format_polygon(region.polygon)) for region in bounding_regions)

def format_polygon(polygon):
    if not polygon:
        return "N/A"
    return ", ".join(["[{}, {}]".format(p.x, p.y) for p in polygon])

def analyze_invoice(url_input, isUrl, filePath):

    if(len(url_input) == 0):
        invoiceUrl = "https://docs.google.com/uc?export=download&id=1BnyEGi00M1vL5mYsViZj2bv5cIIKDRZN"
        invoiceUrl2 = "https://raw.githubusercontent.com/Azure-Samples/cognitive-services-REST-api-samples/master/curl/form-recognizer/sample-invoice.pdf"
        invoiceUrl3 = "https://docs.google.com/uc?export=download&id=1g6kL0DxdjsK-J-c5d-iZZp9IyQlWXHid"
    else:
        invoiceUrl = url_input
    document_analysis_client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(key))

    output_file = "C:\\Users\\razimi\\OneDrive - Walter Surface Technologies\\Desktop\\python\\form_recognizer\\output_prebuilt.csv"
    file_exists = os.path.exists(output_file)


    with open(output_file, "a", newline="") as csv_file:
        writer = csv.writer(csv_file)

        # if not file_exists:
            # writer.writerow(["--------Recognizing invoice--------"])

        if(isUrl):
            poller = document_analysis_client.begin_analyze_document_from_url("prebuilt-invoice", invoiceUrl)
            invoices = poller.result()

        else:
            with open(filePath, "rb") as f:
                poller = document_analysis_client.begin_analyze_document("prebuilt-invoice", document=f, locale="en-US")
                invoices = poller.result()

        for idx, invoice in enumerate(invoices.documents):
            writer.writerow(["--------Recognizing invoice--------"])

            # Write field values and confidence scores to the CSV file
            fields = [
                ("Customer Name", invoice.fields.get("VendorName")),
                ("Customer Address", invoice.fields.get("VendorAddress")),
                # ("Customer Address Recipient", invoice.fields.get("VendorAddressRecipient")),
                ("Supplier Name", invoice.fields.get("CustomerName")),
                ("Supplier Id", invoice.fields.get("CustomerId")),
                ("Supplier Address", invoice.fields.get("CustomerAddress")),
                # ("Supplier Address Recipient", invoice.fields.get("CustomerAddressRecipient")),
                ("PO Number", invoice.fields.get("InvoiceId")),
                ("PO Date", invoice.fields.get("InvoiceDate")),
                ("Payment Term", invoice.fields.get("PaymentTerm")),
                ("Due Date", invoice.fields.get("DueDate")),
                ("Purchase Order", invoice.fields.get("PurchaseOrder")),
                ("Billing Address", invoice.fields.get("BillingAddress")),
                ("Billing Address Recipient", invoice.fields.get("BillingAddressRecipient")),
                ("Deliver To", invoice.fields.get("ShippingAddress")),
                ("Deliver To Recipient", invoice.fields.get("ShippingAddressRecipient")),
                ("Net Total", invoice.fields.get("SubTotal")),
                ("PO Total", invoice.fields.get("InvoiceTotal")),
                ("Service Start Date", invoice.fields.get("ServiceStartDate")),
                ("Service End Date", invoice.fields.get("ServiceEndDate")),
                # ("Total Tax", invoice.fields.get("TotalTax")),
                # ("Previous Unpaid Balance", invoice.fields.get("PreviousUnpaidBalance")),
                # ("Amount Due", invoice.fields.get("AmountDue")),
                # ("Service Address", invoice.fields.get("ServiceAddress")),
                # ("Service Address Recipient", invoice.fields.get("ServiceAddressRecipient")),
                # ("Remittance Address", invoice.fields.get("RemittanceAddress")),
                # ("Remittance Address Recipient", invoice.fields.get("RemittanceAddressRecipient"))
            ]

            for field_name, field in fields:
                if field:
                    writer.writerow([field_name, field.value])
                #else:
                    #writer.writerow([field_name, "N/A"])

            writer.writerow(["Invoice items:"])
            for idx, item in enumerate(invoice.fields.get("Items").value):
                writer.writerow(["Item #{}".format(idx + 1)])
                item_description = item.value.get("Description")
                item_quantity = item.value.get("Quantity")
                item_unit = item.value.get("Unit")
                item_unit_price = item.value.get("UnitPrice")
                item_product_code = item.value.get("ProductCode")
                item_date = item.value.get("Date")
                item_tax = item.value.get("Tax")
                item_amount = item.value.get("Amount")

                if item_product_code:
                    writer.writerow(["Product", item_product_code.value])

                if item_description:
                    writer.writerow(["Description", item_description.value])

                if item_unit:
                    writer.writerow(["U/M", item_unit.value])

                if item_quantity:
                    writer.writerow(["Order QTY", item_quantity.value])
                else:
                    writer.writerow(["Order QTY", (item_amount.value.amount / item_unit_price.value.amount)])

                if item_unit_price:
                    writer.writerow(["Price", item_unit_price.value])

                if item_date:
                    writer.writerow(["Date", item_date.value])
                if item_tax:
                    writer.writerow(["Tax", item_tax.value])
                if item_amount:
                    writer.writerow(["Extension", item_amount.value])

            writer.writerow([])  # Write an empty row to separate invoices
            writer.writerow(["----------------------------------------"])

# this method opens both the invoice (pdf) and the output (csv) at the same time
def open_files_side_by_side(pdf_path, excel_path):
    # Open the PDF file using the default system viewer
    subprocess.Popen([pdf_path], shell=True)

    # Open the Excel file using the default system application
    subprocess.Popen([excel_path], shell=True)

    # Wait for the applications to open
    pyautogui.sleep(2)

    # Get the screen dimensions
    screen_width, screen_height = pyautogui.size()

    # Set the PDF window position
    pdf_window_x = 0
    pdf_window_y = 0
    pdf_window_width = screen_width // 2
    pdf_window_height = screen_height

    # Set the Excel window position
    excel_window_x = screen_width // 2
    excel_window_y = 0
    excel_window_width = screen_width // 2
    excel_window_height = screen_height

    # Move and resize the PDF window
    pyautogui.getWindowsWithTitle('PDF Viewer')[0].resizeTo(pdf_window_width, pdf_window_height)
    pyautogui.getWindowsWithTitle('PDF Viewer')[0].moveTo(pdf_window_x, pdf_window_y)

    # Move and resize the Excel window
    pyautogui.getWindowsWithTitle('Microsoft Excel')[0].resizeTo(excel_window_width, excel_window_height)
    pyautogui.getWindowsWithTitle('Microsoft Excel')[0].moveTo(excel_window_x, excel_window_y)

# main
# Start the main loop
window.mainloop()
