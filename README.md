# MDB to XLSX Converter

This is a Python-based GUI tool that converts Microsoft Access (`.mdb`) database files to Excel (`.xlsx`) files. The tool allows users to easily select an `.mdb` file, choose an output location, and convert the database tables into separate sheets within an Excel file.

## Features
- **User-friendly GUI**: The tool uses a simple interface built with `tkinter`.
- **Multiple Tables**: Converts all tables within the `.mdb` file to separate sheets in the Excel file.
- **Progress Bar**: A progress bar shows the conversion progress, giving users feedback during the process.

## Prerequisites
Make sure you have Python installed on your system. Additionally, install the required Python packages using pip:

```bash
pip install pandas pyodbc openpyxl
Requirements

    Python 3.x
    pandas
    pyodbc
    openpyxl
    Microsoft Access Database Engine: Make sure you have the necessary ODBC drivers installed for Microsoft Access. You may need to install the Microsoft Access Database Engine.

Usage

    Clone or download the repository.


Navigate to the directory where the script is located.

Run the script using Python:

bash

    python main.py

    Select your .mdb file using the "Browse..." button in the interface.

    Choose a location to save the .xlsx file.

    Click "Convert". The tool will start the conversion process and display the progress.

    Once the process is completed, a success message will be shown.

Known Issues

    If you encounter an error related to the Access ODBC driver, make sure that the correct driver is installed on your system.
    The tool currently supports only .mdb files. For .accdb files, additional modifications may be required.

License

This project is licensed under the MIT License - see the LICENSE file for details.
Contributing

Contributions are welcome! Feel free to submit a pull request or open an issue if you encounter any problems.
Acknowledgments

    This tool was developed using Python and the pandas, pyodbc, and openpyxl libraries.
