# Excel Consolidator to SQLite Streamlit App

This Streamlit application allows users to upload multiple Excel files (`.xlsx` and/or `.xlsb`), define data types for each column, and consolidate them into a single dataset. The application performs data cleaning, ensures column consistency, and provides the consolidated data for preview and download as an SQLite database file.

## Features

*   **Multiple File Upload:** Upload one or more Excel files (`.xlsx`, `.xlsb`).
*   **Dynamic Column Typing:** Define data types (Text, Numeric, Date) for each column based on the first valid uploaded file.
*   **Data Cleaning:**
    *   Handles various representations of missing data (empty strings, "N/A", `NaN`) and converts them to a consistent `pd.NA`.
    *   Removes trailing empty rows from each sheet.
*   **Data Transformation:**
    *   Adds a `Source_File` column to track the origin of each row.
    *   Formats date columns to `YYYY-MM-DD` string representation.
*   **Column Consistency:** Enforces that all processed files adhere to the column structure (name and order) of the first valid file uploaded. Files with mismatched structures are skipped with a warning.
*   **Data Preview:** Displays the first 100 rows of the consolidated data.
*   **SQLite Export:** Allows downloading the consolidated data as a binary SQLite database file (`consolidated_data.db`).
*   **User Guidance:** Includes warnings about maintaining consistent column structures in uploaded files.
*   **Performance:** Utilizes Streamlit's caching (`@st.cache_data`) for faster reprocessing and specifies `dtype='object'` during initial Excel file reads for improved loading speed.

## Requirements

*   Python 3.7+
*   The following Python libraries:
    *   `streamlit`
    *   `pandas`
    *   `openpyxl` (for `.xlsx` files)
    *   `pyxlsb` (for `.xlsb` files)
    *   `sqlite3` (usually part of standard Python library)

## How to Run

1.  **Clone the repository or download the files.**
2.  **Install dependencies:**
    Open your terminal or command prompt and navigate to the project directory. Then run:
    ```bash
    pip install streamlit pandas openpyxl pyxlsb
    ```
3.  **Run the Streamlit application:**
    In the same directory, execute:
    ```bash
    streamlit run excel_consolidator.py
    ```
    (Assuming your main application file is named `excel_consolidator.py`)

    The application should open in your default web browser.

## Usage Instructions

1.  **Upload Files:**
    *   Use the file uploader to select one or more Excel files (`.xlsx` or `.xlsb`).
    *   A warning message reminds you that all files should ideally have the same column structure.
2.  **Step 1: Define Column Data Types:**
    *   Once files are uploaded, the application will read the column headers from the first valid file.
    *   For each column, select the appropriate data type:
        *   **Text:** For string data.
        *   **Numeric:** For numerical data (integers or decimals). Values that cannot be converted to numeric will become `NA`.
        *   **Date:** For date values. Values that cannot be converted to dates will become `NA`. Dates will be standardized to `YYYY-MM-DD` format.
    *   The application attempts to suggest a type based on common keywords in column names.
3.  **Step 2: Process Files & Preview:**
    *   Click the "Process Files & Preview" button.
    *   The application will:
        *   Read each uploaded file.
        *   Validate its column structure against the standard set by the first file.
        *   Clean and prepare the data according to the defined types.
        *   Remove trailing empty rows.
        *   Add the `Source_File` column.
        *   Concatenate all processed data.
    *   A progress bar will show the status.
    *   A preview of the first 100 rows of the consolidated data will be displayed.
4.  **Step 3: Download SQLite Database:**
    *   If processing is successful and data is available, a download button will appear.
    *   Click "Download consolidated_data.db" to save the SQLite database file containing all the consolidated data in a table named `DATA`.

## Important Notes

*   **Column Consistency is Crucial:** The application uses the column names and order from the *first successfully read, non-empty Excel file* as the standard. Subsequent files *must* have the exact same column names and order. Files that do not match this structure will be skipped, and a warning will be issued.
*   **Data Type Coercion:** If data in a column cannot be converted to the specified Numeric or Date type, it will be replaced with `NA` (Not Available).
*   **File Encoding:** Assumes standard Excel file encodings.
*   **Large Files:** While optimized, processing very large Excel files can consume significant memory and time. Streamlit's default file upload limit (200MB per file if running locally, potentially different if deployed) also applies.
*   **Error Handling:** The application includes error messages for common issues, such as problems reading files or during data processing.

---
Streamlit Excel Consolidator
