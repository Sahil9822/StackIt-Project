import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Setting up Google Sheets API and creating client
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credential = ServiceAccountCredentials.from_json_keyfile_name('stackit-400715-04c704f6be17.json', scope)
client = gspread.authorize(credential)

# Global variable for the column selection window
col_selection__win = None

# Global variable to store selected columns
cols_import = []

# Global variable to store the CSV file path
csv_file_path = None

# Function to handle CSV import
def import_csv():
    global col_selection__win  # Make column_selection_window a global variable
    global cols_import  # Make columns_to_import a global variable
    global csv_file_path  # Make csv_file_path a global variable

    # Open a file dialog to allow the user to select a CSV file
    csv_file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])

    if not csv_file_path:
        return

    try:
        # Create a dialog to let the user select columns for import
        col_selection__win = tk.Toplevel(root)
        col_selection__win.title("Column Selection")

        col_sel_frame = tk.Frame(col_selection__win)
        col_sel_frame.pack()

        # Read the CSV file in chunks to reduce memory usage
        chunk_size = 10**6  # Adjust the chunk size as needed
        for chunk in pd.read_csv(csv_file_path, chunksize=chunk_size):
            for col_name in chunk.columns:
                tk.Checkbutton(col_sel_frame, text=col_name, command=lambda col=col_name: add_column(col)).pack()

        tk.Button(col_sel_frame, text="Import", command=import_data).pack()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while reading the CSV file: {str(e)}")
        return

# Function to add columns to the list of columns to import
def add_column(col_name):
    global cols_import
    cols_import.append(col_name)

# Function to import data into Google Sheets
def import_data():
    global col_selection__win  # Make column_selection_window a global variable
    global cols_import  # Make columns_to_import a global variable
    global csv_file_path  # Make csv_file_path a global variable

    # Close the column selection window
    if col_selection__win:
        col_selection__win.withdraw()

    if not cols_import:
        messagebox.showerror("Error", "Please select at least one column to import.")
        return

    if not csv_file_path:
        messagebox.showerror("Error", "No CSV file selected.")
        return

    # Open the Google Sheet
    
    sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1bs4zOAmXP5D-cwl4hdKI-NWZ1XNXJZpDCKrH7Y0KuPE/edit?usp=sharing')
    worksheet = sheet.worksheet('StackIt')

    # Clear the existing data in the Google Sheet
    worksheet.clear()

    # Read the CSV file again in chunks and import selected columns
    chunk_size = 10**6  # Same chunk size as during selection
    for chunk in pd.read_csv(csv_file_path, chunksize=chunk_size):
        filtered_data = chunk[cols_import]

        # Convert the filtered data to a list of lists
        data_to_import = [filtered_data.columns.tolist()] + filtered_data.values.tolist()

        # Append the filtered data to the Google Sheet
        worksheet.append_rows(data_to_import)

    messagebox.showinfo("Success", "Data imported successfully to Google Sheets!")
# Create the main application window
root = tk.Tk()
root.title("CSV Importer")

# Create a button to trigger CSV import
import_button = tk.Button(root, text="Import CSV", command=import_csv)
import_button.pack()

root.mainloop()