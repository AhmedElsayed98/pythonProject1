import tkinter as tk
from tkinter import filedialog
import pyodbc
import pandas as pd

# Function to open the Microsoft Access database and store the connection and cursor
def open_access_database():
    global access_conn, access_cursor

    db_file_path = filedialog.askopenfilename(title="Select Microsoft Access Database File", filetypes=[("Access Database Files", "*.mdb")])

    if db_file_path:
        # Set up the ODBC connection
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            rf"DBQ={db_file_path};"
        )
        access_conn = pyodbc.connect(conn_str)
        access_cursor = access_conn.cursor()
        print("Connected to the Access database:", db_file_path)

# Function to open the Excel sheet (template) and store it in a Pandas DataFrame
def open_excel_sheet():
    global excel_df, excel_file_path

    excel_file_path = filedialog.askopenfilename(title="Select Excel Template", filetypes=[("Excel Files", "*.xlsx")])

    if excel_file_path:
        excel_df = pd.read_excel(excel_file_path)
        print("Loaded Excel template:", excel_file_path)

# Create the main GUI window
root = tk.Tk()
root.title('File Import Tool')

# Create a button to open the Microsoft Access database
access_button = tk.Button(root, text="Open Access Database", command=open_access_database)
access_button.pack()

# Create a button to open the Excel sheet (template)
excel_button = tk.Button(root, text="Open Excel Template", command=open_excel_sheet)
excel_button.pack()

# Initialize global variables
access_conn = None
access_cursor = None
excel_df = None
excel_file_path = None

# Start the GUI application
root.mainloop()
print(type(excel_df))
print(type(excel_file_path))
# You can now access the 'access_conn', 'access_cursor', 'excel_df', and 'excel_file_path' outside the functions.
