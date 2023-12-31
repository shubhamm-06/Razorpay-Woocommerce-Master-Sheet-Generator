import requests
import argparse
from datetime import datetime
import razorpay
import gspread
import pandas as pd
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run_flow
from oauth2client.file import Storage

# Common start and end dates
common_start_date = '2023-09-01'
common_end_date = '2023-09-07'

#Google API Authentication
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# OAuth 2.0 credentials for installed applications
flow = OAuth2WebServerFlow(
    client_id="",
    client_secret="",
    scope=scope,
    redirect_uri="http://localhost"
)

# Try to get valid credentials, either from the file or by running the flow
credentials = Storage('credentials.dat').get()

if credentials is None or credentials.invalid:
    flags = argparse.Namespace()
    credentials = run_flow(flow, Storage('credentials.dat'), flags)

# Authorize the credentials
gc = gspread.authorize(credentials)


# WooCommerce API details
woo_url = 'https://wisdomhatch.com'
woo_consumer_key = ''
woo_consumer_secret = ''

# Razorpay API details
razorpay_api_key = ""
razorpay_api_secret = ""

# WooCommerce API endpoint for orders
woo_endpoint = f'{woo_url}/wp-json/wc/v3/orders'

# Request parameters for WooCommerce
woo_params = {
    'consumer_key': woo_consumer_key,
    'consumer_secret': woo_consumer_secret,
    'after': datetime.strptime(common_start_date, '%Y-%m-%d').isoformat(),
    'before': datetime.strptime(common_end_date, '%Y-%m-%d').isoformat(),
    "per_page": 100  # Increase this if you expect many orders
}
print("Credentials Validated \n Colleting Woocommerce data")
# Make the WooCommerce API request
woo_response = requests.get(woo_endpoint, params=woo_params)

# Check if the WooCommerce request was successful (status code 200)
if woo_response.status_code == 200:
    # Convert the JSON response to a DataFrame
    woo_orders_data = woo_response.json()
    woo_orders_df = pd.json_normalize(woo_orders_data)
    print("Woo Data Collected \n Colleting Razorpay data")

else:
    print(f'Error: {woo_response.status_code} - {woo_response.text}')

# Razorpay API endpoint for payments
razorpay_endpoint = 'https://api.razorpay.com/v1/payments'

# Set the query parameters for Razorpay
razorpay_params = {
    "from": int(datetime.strptime(common_start_date, '%Y-%m-%d').timestamp()),  # Convert to integer
    "to": int(datetime.strptime(common_end_date, '%Y-%m-%d').timestamp()),      # Convert to integer
    "count": 100,  # Specify the number of payments to be fetched
}

# Create a Razorpay client
razorpay_client = razorpay.Client(auth=(razorpay_api_key, razorpay_api_secret))

# Get the list of payments (orders in Razorpay are referred to as payments)
razorpay_response = razorpay_client.payment.all(data=razorpay_params)

# Check the Razorpay response status code
if razorpay_response.get('error_code') is None:
    # Extract the data from the Razorpay response
    razorpay_data = razorpay_response['items']

    # Convert Razorpay data to DataFrame
    razorpay_df = pd.json_normalize(razorpay_data)
    print("Razorpay Data Collected \n Converting Json files into excel and Merging it.")
else:
    print(f"Error: Unable to get Razorpay order data. {razorpay_response['error_description']}")

# Keep only numeric values in the 'description' column in Razorpay sheet
razorpay_df['description'] = razorpay_df['description'].str.replace(r'\D', '', regex=True)

# Convert 'description' column to string in both DataFrames
woo_orders_df['id'] = woo_orders_df['id'].astype(str)
razorpay_df['description'] = razorpay_df['description'].astype(str)

# Merge dataframes based on a common key (order_id in this case)
master_df = pd.merge(razorpay_df, woo_orders_df,  how='left', left_on='description', right_on='id', suffixes=('_woo', '_razorpay'))

# Drop rows where 'description' is null
master_df = master_df.dropna(subset=['description'])

# Convert 'captured' column to string in master_df
master_df['captured'] = master_df['captured'].astype(str)

# Filter based on lowercase comparison
master_df = master_df[master_df['captured'].str.lower() == 'true']

columns_to_keep = ['amount','captured','description','email','date_created' ,'billing.first_name','billing.last_name','billing.city','billing.country','billing.email','billing.phone']

# Create or open a Google Sheet
sheet_name = 'new 2'
try:
    # Try to open the sheet if it already exists
    sheet = gc.open(sheet_name)
except gspread.exceptions.SpreadsheetNotFound:
    # If the sheet doesn't exist, create a new one
    sheet = gc.create(sheet_name)

# Select the first (and only) worksheet in the spreadsheet
worksheet = sheet.sheet1

# Write headers to the sheet if it's a new sheet
if worksheet.row_values(1) == []:
    headers = ['Amount', 'Captured', 'Description', 'Email', 'Date Created', 'First Name', 'Last Name', 'City', 'Country', 'Billing Email', 'Phone']
    worksheet.append_row(headers)

# Get existing descriptions in the sheet
existing_descriptions = worksheet.row_values(1)

# Check if 'Description' is in the headers
if 'Description' in existing_descriptions:
    # Get the index of 'Description' column
    description_index = existing_descriptions.index('Description') + 1

    # Get existing descriptions in the sheet (excluding the header)
    existing_descriptions = worksheet.col_values(description_index)[1:]

    # Iterate through the processed data and append each row to the sheet
    for index, row in master_df[columns_to_keep].iterrows():
        # Check for duplicates based on the 'Description' column
        if row['description'] not in existing_descriptions:
            # Convert any float values to strings
            row_str = [str(value) if isinstance(value, float) else value for value in row.tolist()]

            # Append each row to the sheet without checking for duplicates
            worksheet.append_row(row_str)

    worksheet = sheet.sheet1

    # Get all values from the sheet
    values = worksheet.get_all_values()

    # Check if 'Date Created' is in the headers
    header_row = values[0]
    try:
        date_created_index = header_row.index("Date Created")
    except ValueError:
        print("Error: Column 'Date Created' not found.")
        exit()

    # Filter out rows where the "Date Created" value is NaN or null
    filtered_values = [row for row in values[1:] if row[date_created_index]]

    # Sort the data by the "Date Created" column
    filtered_values.sort(key=lambda row: row[date_created_index], reverse=True)

    # Clear the existing data in the sheet
    worksheet.clear()

    # Write headers to the sheet
    worksheet.append_row(header_row)

    # Write the sorted and filtered data to the sheet
    worksheet.append_rows(filtered_values)

    print(f'Data successfully written and Sorted to the Google Sheet: {sheet_name}')
else:
    print("Error: 'Description' column not found in the sheet headers.")