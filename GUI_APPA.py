import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
import pyodbc
# Global variables to store connection, DataFrame, and file paths
conn = None
df = None
access_db_file_path = None
excel_file_path = None
cursor = None
def open_access_database():
    global conn
    global access_db_file_path
    global cursor

    db_file_path = filedialog.askopenfilename(title="Select Microsoft Access Database File", filetypes=[("Access Database Files", "*.mdb")])

    if db_file_path:
        # Set up the ODBC connection
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            rf"DBQ={db_file_path};"
        )
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        print("Connected to the Access database:", db_file_path)
        access_db_file_path = db_file_path
def open_excel_sheet():
    global df
    global excel_file_path

    excel_file_path = filedialog.askopenfilename(title="Select Excel Sheet", filetypes=[("Excel Files", "*.xlsx")])

    if excel_file_path:
        # Read the Excel sheet into a Pandas DataFrame
        df = pd.read_excel(excel_file_path)
        print("Opened Excel sheet:", excel_file_path)
def process_data():
    # Add your data processing logic here

    wbts_df = pd.read_excel(excel_file_path)
    wbts2 = int(wbts_df.iloc[0]['New_WBTS'])
    old_cids = wbts_df['Old_Cids'].to_list()
    new_cids = wbts_df['New_Cids'].to_list()
    new_wbts = wbts_df['New_WBTS'].to_list()
    old_wbts = wbts_df['Old_WBTS'].to_list()
    # rnc_df = pd.read_excel(wbts_file_path)
    rncid = int(wbts_df.iloc[0]['RNCID'])

    print(old_cids)
    print(new_cids)

    value_mapping = dict(zip(old_cids, new_cids))
    value_mapping2 = dict(zip(old_wbts, new_wbts))
    # Specify the path for the new folder
    current_dir = os.path.dirname(excel_file_path)
    create_dic = "create_CSVs"
    delete_dic = "delete_CSVs"
    create_dir_folder_path = os.path.join(current_dir, create_dic)  # Current working directory
    del_dir_folder_path = os.path.join(current_dir, delete_dic)  # Current working directory
    # rnc_df = pd.read_excel(wbts_file_path)
    # rncid = int(wbts_file_path.iloc[0]['RNCID'])

    # Check if the folder already exists
    if not os.path.exists(create_dir_folder_path):
        # Create the new folder
        os.mkdir(create_dir_folder_path)
        print(f"Folder '{create_dir_folder_path}' created successfully at {current_dir}")
    else:
        print(f"Folder '{create_dir_folder_path}' already exists at {current_dir}")
    # Check if the folder already exists
    if not os.path.exists(del_dir_folder_path):
        # Create the new folder
        os.mkdir(del_dir_folder_path)
        print(f"Folder '{del_dir_folder_path}' created successfully at {current_dir}")
    else:
        print(f"Folder '{del_dir_folder_path}' already exists at {current_dir}")

    del_col = 'TargetCellDN'
    del_col2 = 'name'

    # Get table names from the database
    tables = [table.table_name for table in cursor.tables(tableType="TABLE")]

    wcel = 'A_WCEL'
    adjci = 'CId'
    # adjs_partition
    create_df = pd.DataFrame()
    selected_table1 = 'A_ADJS'
    # Read the first table into a dataframe where WBTSId column has wbts1 values and RncId has the value of rncid
    query1 = f"SELECT * FROM [{selected_table1}] WHERE RncId = ?"
    df1 = pd.read_sql(query1, conn, params=(rncid,))
    pref = 'A_'
    title1 = selected_table1.title()
    if title1.startswith(pref):
        title = title1[len(pref):].upper()
        title2 = title1[len(pref):]
    title.upper()
    col_string = str(title) + 'Id'
    if selected_table1.title() == 'A_ADJW':
        col_string2 = str(title2) + 'CId'
    else:
        col_string2 = str(title2) + 'CI'
    if col_string in df1.columns:
        print('yes')
    else:
        print(df1.columns[: 10])
    key_col = df1.columns.get_loc(col_string)

    create_df = df1.copy()
    create_df_1 = create_df[create_df[col_string2].isin(old_cids)]
    create_df_2 = create_df[create_df['WBTSId'].isin(old_wbts)]
    create_df_final = pd.concat([create_df_2, create_df_1], ignore_index=True)
    create_df_final['WBTSId'] = create_df_final['WBTSId'].replace(value_mapping2)
    create_df_final[col_string2] = create_df_final[col_string2].replace(value_mapping)
    query3 = f"Select * FROM [{selected_table1}] WHERE RncId = ?"
    adj_df = pd.read_sql(query3, conn, params=(rncid,))
    delete_df = pd.DataFrame()
    delete_df = adj_df[adj_df['WBTSId'].isin(old_wbts)]
    df_src_d = delete_df.iloc[:, :key_col + 1]
    df_tar_d = pd.DataFrame()
    df_tar_d = df1[(df1[col_string2].isin(old_cids))]
    df_tar_d2 = df_tar_d.iloc[:, :key_col + 1]
    operation = 'delete'
    final_delete = pd.DataFrame()
    final_delete = pd.concat([df_src_d, df_tar_d2], ignore_index=True)
    final_delete['$Operation'] = operation
    # Prompt user to choose the location to save the CSV file
    # csv_file_path = filedialog.asksaveasfilename(title="Save Concatenated Dataframe as CSV", defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    path_del_ajds = os.path.join(del_dir_folder_path, 'A_ADJS_DELETE.csv')
    final_delete.to_csv(path_del_ajds, index=False)
    # Save the concatenated dataframe to a CSV file
    path_create_ajds = os.path.join(create_dir_folder_path, 'A_ADJS_Create.csv')
    del_col11 = create_df_final.columns.get_loc(del_col)
    create_df_final = create_df_final.iloc[:, :del_col11]
    create_df_final.to_csv(path_create_ajds, index=False)
    # adji partition
    selected_table2 = 'A_ADJI'
    query2 = f"SELECT * FROM [{selected_table2}] WHERE RncId = ?"
    df_adji1 = pd.read_sql(query2, conn, params=(rncid,))
    pref = 'A_'
    title1 = selected_table2.title()
    if title1.startswith(pref):
        title = title1[len(pref):].upper()
        title2 = title1[len(pref):]
    title.upper()
    col_string = str(title) + 'Id'
    if selected_table2.title() == 'A_ADJW':
        col_string2 = str(title2) + 'CId'
    else:
        col_string2 = str(title2) + 'CI'
    if col_string in df_adji1.columns:
        print('yes')
    else:
        print(df_adji1.columns[: 10])
    key_col = df_adji1.columns.get_loc(col_string)

    create_df_adji = df_adji1.copy()
    create_df_adji2 = create_df_adji[create_df_adji[col_string2].isin(old_cids)]
    create_df_adji3 = create_df_adji[create_df_adji['WBTSId'].isin(old_wbts)]
    create_df_adji_final = pd.concat([create_df_adji2, create_df_adji3], ignore_index=True)
    create_df_adji_final['WBTSId'] = create_df_adji_final['WBTSId'].replace(value_mapping2)
    create_df_adji_final[col_string2] = create_df_adji_final[col_string2].replace(value_mapping)

    query4 = f"Select * FROM [{selected_table2}] WHERE RncId = ?"
    adji_df = pd.read_sql(query4, conn, params=(rncid,))
    delete_adji_df = pd.DataFrame()
    delete_adji_df = adji_df[adji_df['WBTSId'].isin(old_wbts)]
    df_src_d_adji = delete_adji_df.iloc[:, :key_col + 1]
    df_tar_d_adji = pd.DataFrame()
    df_tar_d_adji = df_adji1[(df_adji1[col_string2].isin(old_cids))]
    df_tar_adji_d2 = df_tar_d_adji.iloc[:, :key_col + 1]
    final_delete_adji = pd.DataFrame()
    final_delete_adji = pd.concat([df_src_d_adji, df_tar_adji_d2], ignore_index=True)
    final_delete_adji['$Operation'] = operation
    path_del_ajdi = os.path.join(del_dir_folder_path, 'A_ADJI_DELETE.csv')
    final_delete_adji.to_csv(path_del_ajdi, index=False)
    # Save the concatenated dataframe to a CSV file
    path_create_ajdi = os.path.join(create_dir_folder_path, 'A_ADJI_Create.csv')
    del_col11 = create_df_adji_final.columns.get_loc(del_col)
    create_df_adji_final = create_df_adji_final.iloc[:, :del_col11]
    create_df_adji_final.to_csv(path_create_ajdi, index=False)

    # adjw partition
    selected_table3 = 'A_ADJW'
    query5 = f"SELECT * FROM [{selected_table3}] WHERE RncId = ?"
    df_adjw1 = pd.read_sql(query5, conn, params=(rncid,))
    pref = 'A_'
    title1 = selected_table3.title()
    if title1.startswith(pref):
        title = title1[len(pref):].upper()
        title2 = title1[len(pref):]
    title.upper()
    col_string = str(title) + 'Id'
    col_string2 = str(title2) + 'CId'
    key_col = df_adjw1.columns.get_loc(col_string)
    create_df_adjwi = df_adjw1.copy()
    create_df_adjw2 = create_df_adjwi[create_df_adjwi[col_string2].isin(old_cids)]
    delete_adjw_df = create_df_adjw2.copy()
    df_tar_adjw_d2 = delete_adjw_df.iloc[:, :key_col + 1]
    create_df_adjw2[col_string2] = create_df_adjw2[col_string2].replace(value_mapping)
    create_df_adjw2['sac'] = create_df_adjw2['sac'].replace(value_mapping)

    del_colw = 'targetCellDN'
    df_tar_adjw_d2['$Operation'] = operation
    path_del_ajdw = os.path.join(del_dir_folder_path, 'A_ADJW_DELETE.csv')
    path_create_ajdw = os.path.join(create_dir_folder_path, 'A_ADJW_Create.csv')
    del_col11 = create_df_adjw2.columns.get_loc(del_colw)
    create_df_adjw2 = create_df_adjw2.iloc[:, :del_col11]
    create_df_adjw2.to_csv(path_create_ajdw, index=False)
    df_tar_adjw_d2.to_csv(path_del_ajdw, index=False)
    # adjg partition
    selected_table4 = 'A_ADJG'
    query6 = f"SELECT * FROM [{selected_table4}]  WHERE RncId = ?"

    df_adjg1 = pd.read_sql(query6, conn, params=(rncid,))
    pref = 'A_'
    title1 = selected_table4.title()
    if title1.startswith(pref):
        title = title1[len(pref):].upper()
        title2 = title1[len(pref):]
    title.upper()
    col_string = str(title) + 'Id'
    col_string2 = str(title2) + 'CId'
    key_col = df_adjg1.columns.get_loc(col_string)
    df_del_adjg = pd.DataFrame()
    df_del_adjg = df_adjg1[df_adjg1['WBTSId'].isin(old_wbts)]
    df_del_adjg = df_del_adjg.iloc[:, :key_col + 1]
    df_del_adjg['$Operation'] = operation
    df_create_adjg = pd.DataFrame()
    df_create_adjg = df_adjg1[df_adjg1['WBTSId'].isin(old_wbts)]
    df_create_adjg['WBTSId'] = df_create_adjg['WBTSId'].replace(value_mapping2)
    path_del_ajdg = os.path.join(del_dir_folder_path, 'A_ADJG_DELETE.csv')
    df_del_adjg.to_csv(path_del_ajdg, index=False)
    del_col11 = df_create_adjg.columns.get_loc(del_col)
    df_create_adjg = df_create_adjg.iloc[:, :del_col11]
    path_cre_ajdg = os.path.join(create_dir_folder_path, 'A_ADJG_Create.csv')
    df_create_adjg.to_csv(path_cre_ajdg, index=False)

    # adjl partition
    selected_table5 = 'A_ADJL'
    query7 = f"SELECT * FROM [{selected_table5}]  WHERE RncId = ?"
    df_adjL1 = pd.read_sql(query7, conn, params=(rncid,))
    pref = 'A_'
    title1 = selected_table5.title()
    if title1.startswith(pref):
        title = title1[len(pref):].upper()
        title2 = title1[len(pref):]
    title.upper()
    col_string = str(title) + 'Id'

    key_col = df_adjL1.columns.get_loc(col_string)
    # del_col22 = df_adjL1.columns.get_loc(del_col2)
    df_adjls = pd.DataFrame()
    df_adjl_s = df_adjL1[df_adjL1['WBTSId'].isin(old_wbts)]
    df_adjl_s['WBTSId'] = df_adjl_s['WBTSId'].replace(value_mapping2)
    path_create_adjl = os.path.join(create_dir_folder_path, 'A_ADJL_Create.csv')
    # del_col11 = df_adjl_s.columns.get_loc(del_col22)
    # df_adjl_s = df_adjl_s.iloc[:, :del_col11]
    df_adjl_s.to_csv(path_create_adjl, index=False)
    # LTE_MRBTS_LNBTS_LNADJW part
    selected_table6 = 'A_LTE_MRBTS_LNBTS_LNADJW'
    query8 = f"SELECT * FROM [{selected_table6}]  WHERE uTargetRncId = ?"
    lnadjw_1 = pd.read_sql(query8, conn, params=(rncid,))
    col_string = 'lnAdjWId'
    col_string2 = 'uTargetCid'
    key_col = lnadjw_1.columns.get_loc(col_string)
    del_df_lnadjwi = pd.DataFrame()
    del_df_lnadjwi = lnadjw_1.copy()
    del_df_lnadjwi = del_df_lnadjwi[del_df_lnadjwi[col_string2].isin(old_cids)]
    del_df_lnadjwi = del_df_lnadjwi.iloc[:, :key_col + 1]
    del_df_lnadjwi['$Operation'] = operation
    path_lnadjw = os.path.join(del_dir_folder_path, 'LTE_MRBTS_LNBTS_LNADJW.csv')
    del_df_lnadjwi.to_csv(path_lnadjw, index=False)
    # lnrelw part
    selected_table7 = 'A_LTE_MRBTS_LNBTS_LNCEL_LNRELW'
    query9 = f"SELECT * FROM [{selected_table7}]  WHERE uTargetRncId = ?"
    lnrelw_1 = pd.read_sql(query9, conn, params=(rncid,))
    col_string = 'lnRelWId'
    col_string2 = 'uTargetCid'
    key_col = lnrelw_1.columns.get_loc(col_string)
    lnrel_target = pd.DataFrame()
    lnrel_target = lnrelw_1.copy()
    lnrel_target = lnrel_target[lnrel_target[col_string2].isin(old_cids)]
    lnrel_target = lnrel_target.iloc[:, :key_col + 1]
    lnrel_target['$Operation'] = operation
    path_lnrelw = os.path.join(del_dir_folder_path, 'LTE_MRBTS_LNBTS_LNCEL_LNRELW.csv')
    lnrel_target.to_csv(path_lnrelw, index=False)

    messagebox.showinfo("Success", "Kindly check the exported CSVs at the destination directory!"
                                   "Made By: Ahmed Elsayed   ")

root = tk.Tk()
root.title('3G Neighbors Relations TOOL'
           ' By Ahmed Elsayed')
root.geometry("600x300")
root.configure(bg='#F0F0F0')
# Set up the ODBC connection

# Button styles
button_style = {'bg': '#4CAF50', 'fg': 'white', 'padx': 10, 'pady': 5, 'bd': 2, 'relief': 'groove', 'font': ('Helvetica', 10, 'bold')}

# Message label
message_label1 = tk.Label(root, text="Welcome to the 3G Neighbors Relations TOOL!", font=('Helvetica', 12), bg='#F0F0F0')
message_label1.pack(pady=10)

message_label2 = tk.Label(root, text="Current Families:LNRELW-LNADJW-ADJS-ADJI-ADJG-ADJW-ADJL ", font=('Helvetica', 13), bg='#F0F0FF')
message_label2.pack(pady=10)

open_db_button = tk.Button(root, text="Read Dump File as Access Database", command=open_access_database, **button_style)
open_db_button.pack(pady=10)

open_excel_button = tk.Button(root, text="Open Template Excel Sheet", command=open_excel_sheet, **button_style)
open_excel_button.pack(pady=10)
process_button = tk.Button(root, text="Process", command=process_data, **button_style)
process_button.pack(pady=10)
root.mainloop()


conn.close()
#print("Process completed successfully!")

