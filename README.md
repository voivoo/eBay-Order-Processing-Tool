# eBay Order Processing Tool
This Python script is a desktop application (GUI) designed to fetch, process, and manage eBay orders. It helps automate the task of extracting order details from the eBay Fulfillment API, filtering out canceled orders, performing specific SKU transformations, and finally, writing the processed data into an Excel file, avoiding duplicates.

Features
GUI Interface: User-friendly graphical interface built with tkinter for easy input of eBay access tokens, order filters, and Excel file paths.

eBay API Integration: Fetches order details directly from the eBay Sell Fulfillment API.

Configurable Filters: Allows users to specify the number of days to look back for orders and set a limit on the number of orders to retrieve.

Automatic Configuration Saving: Saves user inputs (except the access token for security reasons) to a JSON file for quick reuse.

Order Filtering: Automatically filters out canceled orders.

SKU Transformation Logic: Implements specific business logic to modify SKUs based on predefined rules (e.g., handling 'NF' suffixes, splitting SKUs like "HLMRXXYY" into two lines if XX and YY differ, and adjusting quantities/prices).

Duplicate Prevention: Compares fetched orders with existing data in the Excel file to prevent writing duplicate entries based on orderId.

Excel Export: Writes processed order data to a specified Excel worksheet, automatically finding the next empty row and formatting dates.

Platform-Independent Configuration: Uses platformdirs to store configuration files in a standard location based on the operating system.

Real-time Information Display: Provides feedback to the user within the GUI about the processing status and any errors encountered.

Requirements
Before running the script, ensure you have the following Python libraries installed:

requests

openpyxl

tkinter (usually comes with Python)

platformdirs

You can install them using pip:

pip install requests openpyxl platformdirs

How to Use
Obtain an eBay Access Token: You'll need a valid eBay access token to use the eBay Fulfillment API. This typically involves setting up an application on the eBay Developers Program website.

Prepare your Excel File: Make sure you have an Excel file (.xlsx format) ready where you want to write the order data. You also need to specify the exact name of the worksheet within that file.

Run the Script:

python eBay Order Processing Tool.py

Fill in the GUI:

eBay access token: Paste your eBay access token into the large text field.

### Input Fields
- **eBay access token**: Paste your eBay access token into the large text field.
- **Order Days**: Number of past days to fetch orders from (e.g., 3 for the last 3 days)
- **Orders Limit**: Maximum number of orders to retrieve (e.g., 100)
- **Target Excel File**: Click "Browse" to select an Excel file or enter the full path
- **Worksheet Name**: Name of the worksheet for data export (e.g., Sheet1 or Orders)

### Processing
1. Click "Start Processing" to begin
2. The application will:
   - Fetch orders from eBay API
   - Process SKUs according to business rules
   - Generate an Excel report
   - Open the report automatically when complete

### Progress Monitoring
The information display area at the bottom shows:
- Processing status
- Success/failure messages
- Any errors encountered

## Project Structure
```
.
├── eBay Order Processing Tool.py  # Main application script
└── README.md                    # This file
```

## Project Overview
A Python tool for processing eBay orders with the following features:
- Fetch order data via eBay Fulfillment API
- Process order SKU transformation logic
- Generate Excel order reports

## Usage Instructions
1. Run `eBay Order Processing Tool.py`
2. Enter your eBay API token in the interface
3. Set the date range and order limit
4. Click the "Start Processing" button
5. The generated Excel file will open automatically upon completion

## Technical Implementation
Main implementation steps:

Configuration Management: Loads and saves user preferences (token excluded for security) to a platform-specific JSON file.

UI Interaction: A tkinter GUI allows users to input necessary parameters.

API Fetching: Uses the eBay Fulfillment API (/sell/fulfillment/v1/order) to retrieve orders within a specified date range and limit. It then fetches detailed information for each order.

Initial Data Processing:

Removes orders with a cancel_status of "CANCELED".

Sorts the remaining orders by creationDate (oldest first).

Truncates creationDate to only include the date part (YYYY-MM-DD) for Excel compatibility.

SKU Transformation Logic:

'NF' Suffix Handling: If an SKU ends with 'NF', it creates two separate entries: one with 'NF' removed from the original SKU, and another with SKU "DWR30" and a price of 0.

Specific Type Handling (HLMR, DR, CL, DBL):

Extracts letters, first two digits, and last two digits from the SKU.

Appends a '0' to 10, 11, or 12 in the digit parts.

If the first two digits match the last two, the quantity is doubled, and the SKU is letters + digits.

If they differ, the order is split into two lines: one with letters + first_two_digits, and another with letters + last_two_digits and a price of 0.

Other SKUs: SKUs not matching these patterns are left unchanged.

Excel Integration:

Duplicate Check: Reads existing order_ids from column H of the target Excel sheet to prevent writing duplicates.

Data Writing: Appends the processed and non-duplicate order data to the next available row in the specified Excel worksheet.

Date Formatting: Formats the 'creationDate' column in Excel to DD.MM.YY.

Important Notes
API Token Security: The eBay access token is critical. This script processes it locally but does not save it to the configuration file. Always handle your API tokens with care.

Error Handling: Basic error handling is included for API requests, file operations, and input validation. More robust error handling might be needed for production use.

SKU Logic Specificity: The SKU transformation rules are highly specific to your described business logic. Ensure these rules align with your requirements.

Excel Structure: The script assumes a specific column mapping for writing data (e.g., order ID in column H, date in column A). Adjust the write_orders_to_excel function if your Excel structure differs.

Contribution
Feel free to fork this repository, suggest improvements, or submit pull requests.

License
This project is open-source and available under the MIT License.
