import pandas as pd
import tkinter as tk
from tkinter import filedialog

def extract_numeric(value):
    # Function to extract numeric part from a string
    numeric_part = ''.join(filter(str.isdigit, str(value)))
    return int(numeric_part) if numeric_part else None

def merge_sheets():
    try:
        # Ask the user to select the razor pay sheet
        raz_path = filedialog.askopenfilename(title="Select Razor Pay Excel File", filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        # Ask the user to select the woo commerce sheet
        woo_path = filedialog.askopenfilename(title="Select Woo Commerce Excel File", filetypes=[("Excel Files", "*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        # Read the data from the selected sheets
        raz_df = pd.read_excel(raz_path)
        woo_df = pd.read_excel(woo_path)

        # Extract numeric part from "invoice_id" column
        raz_df["invoice_id"] = raz_df["invoice_id"].apply(extract_numeric)

        # Merge the sheets based on the common key
        merged_df = pd.merge(raz_df, woo_df, left_on="invoice_id", right_on="Order ID", how="outer")

        # Handle conflicts by prioritizing razor pay sheet
        merged_df = merged_df.fillna(raz_df)

        # Keep only the specified columns
        columns_to_keep = ['created_at', 'Full Name (Billing)', 'State Name (Billing)', 'Phone (Billing)', 'Email (Billing)', 'Order Total Amount', 'invoice_id']
        merged_df = merged_df[columns_to_keep]

        # Ask the user to choose the download location for the master sheet
        download_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx;*.xls")])

        # Save the merged sheet to the selected download location
        merged_df.to_excel(download_path, index=False)

        # Inform the user that the merge is successful
        result_label.config(text="Master sheet made Successfully!\nMaster sheet saved at:\n" + download_path, relief=tk.SUNKEN, bd=3)

    except Exception as e:
        result_label.config(text="Error: " + str(e), fg="red")

# Create the main GUI window
root = tk.Tk()
root.title("Excel Sheet Merger")
root.configure(bg="#1E1E1E")  # Set background color to dark gray

# Set the window size
window_width = 600
window_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width / 2) - (window_width / 2)
y_coordinate = -10  # Place the window at the top
root.geometry("%dx%d+%d+%d" % (window_width, window_height, x_coordinate, y_coordinate))


# Create a button to trigger the merging process
merge_button = tk.Button(root, text="MERGE SHEETS", command=merge_sheets, bg="#4CAF50", fg="white", padx=10, pady=5, font=("Courier", 18))
merge_button.pack(pady=10, side=tk.TOP, anchor=tk.N)

# Steps label
steps_label = tk.Label(root, text="Steps:\n1. Click 'Merge Sheets' button.\n2. Select Razor Pay Excel/CSV file.\n3. Select Woo Commerce Excel/CSV file.\n4. Choose the download location for the master sheet.", bg="#1E1E1E", fg="white", font=("Courier", 14), justify="left")
steps_label.pack()

# Exit button
exit_button = tk.Button(root, text="Exit", command=root.destroy, bg="#FF5733", fg="white", padx=10, pady=5, font=("Courier", 16))
exit_button.pack(pady=10, side=tk.TOP, anchor=tk.N)

# Create a label to display the result
result_label = tk.Label(root, text="", pady=10, bg="#1E1E1E", fg="white", font=("Courier", 16), justify="left")
result_label.pack(fill=tk.BOTH, expand=True)


# Start the GUI event loop
root.mainloop()