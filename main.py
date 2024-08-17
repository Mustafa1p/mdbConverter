import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyodbc

# Function to convert .mdb to .xlsx with progress bar
def mdb_to_xlsx(mdb_file_path, output_excel_file, progress_bar):
    try:
        # Create connection to the .mdb file
        conn_str = (
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=" + mdb_file_path + ";"
        )
        conn = pyodbc.connect(conn_str)
        
        # Fetch all table names from the .mdb file
        cursor = conn.cursor()
        table_names = [row.table_name for row in cursor.tables(tableType='TABLE')]

        total_tables = len(table_names)

        # Create a Pandas Excel writer to write data to Excel
        with pd.ExcelWriter(output_excel_file, engine='openpyxl') as writer:
            for index, table_name in enumerate(table_names):
                # Query the table and fetch the data into a DataFrame
                query = f"SELECT * FROM [{table_name}]"
                df = pd.read_sql(query, conn)
                # Write each table to a different sheet in the Excel file
                df.to_excel(writer, sheet_name=table_name, index=False)
                
                # Update progress bar
                progress_bar['value'] = (index + 1) / total_tables * 100
                progress_bar.update_idletasks()

        # Close the connection
        conn.close()
        messagebox.showinfo("Success", "Conversion completed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to select the .mdb file
def select_mdb_file():
    file_path = filedialog.askopenfilename(
        title="Select .mdb File", 
        filetypes=(("Access Database Files", "*.mdb"), ("All Files", "*.*"))
    )
    if file_path:
        mdb_entry.delete(0, tk.END)
        mdb_entry.insert(0, file_path)

# Function to select the output .xlsx file
def select_output_file():
    file_path = filedialog.asksaveasfilename(
        title="Save as Excel File",
        defaultextension=".xlsx",
        filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
    )
    if file_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, file_path)

# Function to trigger the conversion with progress bar
def convert_file():
    mdb_file = mdb_entry.get()
    output_file = output_entry.get()
    if not mdb_file or not output_file:
        messagebox.showwarning("Input Error", "Please select both input and output files.")
    else:
        progress_bar['value'] = 0
        mdb_to_xlsx(mdb_file, output_file, progress_bar)

# Create the main window
root = tk.Tk()
root.title("MDB to XLSX Converter")

# Create and place widgets
tk.Label(root, text="Select .mdb File:").grid(row=0, column=0, padx=10, pady=10)
mdb_entry = tk.Entry(root, width=50)
mdb_entry.grid(row=0, column=1, padx=10, pady=10)
mdb_button = tk.Button(root, text="Browse...", command=select_mdb_file)
mdb_button.grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Save as .xlsx File:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
output_button = tk.Button(root, text="Browse...", command=select_output_file)
output_button.grid(row=1, column=2, padx=10, pady=10)

# Progress bar widget
progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=2, column=0, columnspan=3, padx=10, pady=20)

convert_button = tk.Button(root, text="Convert", command=convert_file)
convert_button.grid(row=3, column=0, columnspan=3, pady=20)

# Start the application
root.mainloop()
