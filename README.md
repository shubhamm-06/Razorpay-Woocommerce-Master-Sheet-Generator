Excel Sheet Merger
Overview
This Python script simplifies the process of combining data from Razor Pay and Woo Commerce into a master sheet. The script offers a graphical interface to make the merging process easy and accessible.

Prerequisites
Before you begin, ensure you have the following:

Python installed on your computer.
The necessary Python libraries, particularly Pandas.
How to Use
Step 1: Run the Script
Execute the script by running the Python file:

bash
Copy code
python excel_sheet_merger.py
Step 2: Merging Sheets
Click the "MERGE SHEETS" button.
Choose your Razor Pay Excel/CSV file when prompted.
Select your Woo Commerce Excel/CSV file when prompted.
Step 3: Choose Download Location
Pick a location to save your newly created master sheet when prompted.
Step 4: View Results
The script will let you know if the merge was successful or display an error message if any issues arise.
Understanding the Code
This script uses Tkinter to create the graphical interface and Pandas for handling the data. Here's a quick look at how the code works:

merge_sheets Function: Manages the merging process.
File Selection: Prompts you to choose your Razor Pay and Woo Commerce files.
Data Processing: Reads and processes your data using Pandas.
Merge Logic: Combines data based on the "invoice_id" and "Order ID" columns.
Conflict Resolution: Resolves any conflicts by prioritizing Razor Pay data.
Column Selection: Keeps only the columns you specify in the master sheet.
Download Location: Asks you to choose where to save the master sheet.
Result Display: Informs you of success or shows an error message.
Additional Notes
This script is designed to make creating a master sheet from Razor Pay and Woo Commerce data straightforward. Feel free to customize the script to fit your specific needs.
