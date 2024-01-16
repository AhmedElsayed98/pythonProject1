import pyodbc
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


# Function to read data from a Microsoft Access database
def read_access_database(database_path, table_name):
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={database_path};'
    connection = pyodbc.connect(conn_str)
    query = f'SELECT * FROM [{table_name}]'
    df = pd.read_sql(query, connection)
    connection.close()
    return df


# Function to browse for a database file
def browse_file(var):
    file_path = filedialog.askopenfilename(filetypes=[("Access Databases", "*.mdb;*.accdb")])
    var.set(file_path)


# Function to handle the "Compare" button click event
def compare_databases():
    # Get the selected file paths
    file_path1 = file_path_var1.get()
    file_path2 = file_path_var2.get()

    # Get the table names
    table_name1 = table_name_var1.get()
    table_name2 = table_name_var2.get()

    try:
        # Prompt the user for the "rnc" and "wbts" values to filter
        rnc_value = simpledialog.askinteger("Enter 'rnc' Value", "Please enter the 'rnc' value:")
        wbts_value = simpledialog.askinteger("Enter 'wbts' Value", "Please enter the 'wbts' value:")

        # Read data from both databases into dataframes
        df1 = read_access_database(file_path1, table_name1)
        df2 = read_access_database(file_path2, table_name2)

        # Filter dataframes based on user-defined 'rnc' and 'wbts' values
        df1_filtered = df1[(df1['RncId'] == rnc_value) & (df1['WBTSId'] == wbts_value)]
        df2_filtered = df2[(df2['RncId'] == rnc_value) & (df2['WBTSId'] == wbts_value)]

        # Create a comparison report Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Comparison Report"
        ws.append(["Column Name", "Difference Found"])

        # Compare dataframes column by column
        differences = {}
        for col in df1_filtered.columns:
            if col in df2_filtered.columns:
                col1_values = df1_filtered[col]
                col2_values = df2_filtered[col]

                # Check for differences in values
                col_diff = (col1_values != col2_values)
                if col_diff.any():
                    differences[col] = col_diff

                    # Write the column name and difference status to the report
                    ws.append([col, "Yes"])
                else:
                    ws.append([col, "No"])

        # Save the comparison report
        report_filename = "Comparison_Report.xlsx"
        wb.save(report_filename)

        # Display a message box with the report filename
        messagebox.showinfo("Comparison Complete", f"Comparison completed. Report saved as {report_filename}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# Create a Tkinter window
window = tk.Tk()
window.title("Access Database Comparator")

# Create and configure widgets
file_label1 = tk.Label(window, text="Select Database 1:")
file_label1.pack()
file_path_var1 = tk.StringVar()
file_path_entry1 = tk.Entry(window, textvariable=file_path_var1)
file_path_entry1.pack()
browse_button1 = tk.Button(window, text="Browse", command=lambda: browse_file(file_path_var1))
browse_button1.pack()

file_label2 = tk.Label(window, text="Select Database 2:")
file_label2.pack()
file_path_var2 = tk.StringVar()
file_path_entry2 = tk.Entry(window, textvariable=file_path_var2)
file_path_entry2.pack()
browse_button2 = tk.Button(window, text="Browse", command=lambda: browse_file(file_path_var2))
browse_button2.pack()

table_name_label1 = tk.Label(window, text="Enter Table Name for Database 1:")
table_name_label1.pack()
table_name_var1 = tk.StringVar()
table_name_entry1 = tk.Entry(window, textvariable=table_name_var1)
table_name_entry1.pack()

table_name_label2 = tk.Label(window, text="Enter Table Name for Database 2:")
table_name_label2.pack()
table_name_var2 = tk.StringVar()
table_name_entry2 = tk.Entry(window, textvariable=table_name_var2)
table_name_entry2.pack()

compare_button = tk.Button(window, text="Compare", command=compare_databases)
compare_button.pack()

# Run the Tkinter main loop
window.mainloop()
