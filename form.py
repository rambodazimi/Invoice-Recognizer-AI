import csv
import os
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# Set the endpoint and key variables with the values from the Azure portal
endpoint = "<Your Endpoint>"
key = "<Your Key>"

def format_bounding_region(bounding_regions):
    if not bounding_regions:
        return "N/A"
    return ", ".join("Page #{}: {}".format(region.page_number, format_polygon(region.polygon)) for region in bounding_regions)

def format_polygon(polygon):
    if not polygon:
        return "N/A"
    return ", ".join(["[{}, {}]".format(p.x, p.y) for p in polygon])

def analyze_invoice(url_input):

    if(len(url_input) == 0):
        invoiceUrl = "https://docs.google.com/uc?export=download&id=1BnyEGi00M1vL5mYsViZj2bv5cIIKDRZN"
        invoiceUrl2 = "https://raw.githubusercontent.com/Azure-Samples/cognitive-services-REST-api-samples/master/curl/form-recognizer/sample-invoice.pdf"
    else:
        invoiceUrl = url_input
    document_analysis_client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(key))

    output_file = "C:\\Users\\razimi\\OneDrive - Walter Surface Technologies\\Desktop\\python\\form_recognizer\\output.csv"
    file_exists = os.path.exists(output_file)

    with open(output_file, "a", newline="") as csv_file:
        writer = csv.writer(csv_file)

        if not file_exists:
            writer.writerow(["--------Recognizing invoice--------"])

        poller = document_analysis_client.begin_analyze_document_from_url("prebuilt-invoice", invoiceUrl)
        invoices = poller.result()

        print("Extracting data...")

        for idx, invoice in enumerate(invoices.documents):
            writer.writerow(["--------Recognizing invoice #{}--------".format(idx + 1)])

            # Write field values and confidence scores to the CSV file
            fields = [
                ("Vendor Name", invoice.fields.get("VendorName")),
                ("Vendor Address", invoice.fields.get("VendorAddress")),
                ("Vendor Address Recipient", invoice.fields.get("VendorAddressRecipient")),
                ("Customer Name", invoice.fields.get("CustomerName")),
                ("Customer Id", invoice.fields.get("CustomerId")),
                ("Customer Address", invoice.fields.get("CustomerAddress")),
                ("Customer Address Recipient", invoice.fields.get("CustomerAddressRecipient")),
                ("Invoice Id", invoice.fields.get("InvoiceId")),
                ("Invoice Date", invoice.fields.get("InvoiceDate")),
                ("Invoice Total", invoice.fields.get("InvoiceTotal")),
                ("Due Date", invoice.fields.get("DueDate")),
                ("Purchase Order", invoice.fields.get("PurchaseOrder")),
                ("Billing Address", invoice.fields.get("BillingAddress")),
                ("Billing Address Recipient", invoice.fields.get("BillingAddressRecipient")),
                ("Shipping Address", invoice.fields.get("ShippingAddress")),
                ("Shipping Address Recipient", invoice.fields.get("ShippingAddressRecipient")),
                ("Sub Total", invoice.fields.get("SubTotal")),
                ("Total Tax", invoice.fields.get("TotalTax")),
                ("Previous Unpaid Balance", invoice.fields.get("PreviousUnpaidBalance")),
                ("Amount Due", invoice.fields.get("AmountDue")),
                ("Service Start Date", invoice.fields.get("ServiceStartDate")),
                ("Service End Date", invoice.fields.get("ServiceEndDate")),
                ("Service Address", invoice.fields.get("ServiceAddress")),
                ("Service Address Recipient", invoice.fields.get("ServiceAddressRecipient")),
                ("Remittance Address", invoice.fields.get("RemittanceAddress")),
                ("Remittance Address Recipient", invoice.fields.get("RemittanceAddressRecipient"))
            ]

            for field_name, field in fields:
                if field:
                    writer.writerow([field_name, field.value])
                else:
                    writer.writerow([field_name, "N/A"])

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

                if item_description:
                    writer.writerow(["Description", item_description.value])
                if item_quantity:
                    writer.writerow(["Quantity", item_quantity.value])
                if item_unit:
                    writer.writerow(["Unit", item_unit.value])
                if item_unit_price:
                    writer.writerow(["Unit Price", item_unit_price.value])
                if item_product_code:
                    writer.writerow(["Product Code", item_product_code.value])
                if item_date:
                    writer.writerow(["Date", item_date.value])
                if item_tax:
                    writer.writerow(["Tax", item_tax.value])
                if item_amount:
                    writer.writerow(["Amount", item_amount.value])

            writer.writerow([])  # Write an empty row to separate invoices
            writer.writerow(["----------------------------------------"])
    print("Exporting the CSV file...")
    print("Data extracted from the document has been saved to: {}".format(output_file))

if __name__ == "__main__":
    url_input = input("Please enter the url of the document:")

    print("Analyzing the document...")

    analyze_invoice(url_input)

    print("100% SUCCESSFUL!")

    input("Press Enter to exit...")

