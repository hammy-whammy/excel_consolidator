import streamlit as st
import pandas as pd
import os
import io
import sqlite3 # For creating SQLite DB
import tempfile # Added for temporary file handling

# --- Cached Helper Function: Load raw DataFrame ---
@st.cache_data
def load_raw_df(file_content_bytes, file_name, dtypes_spec):
    """Loads a DataFrame from Excel file bytes, using specified dtypes."""
    if file_name.endswith('.xlsb'):
        return pd.read_excel(io.BytesIO(file_content_bytes), engine='pyxlsb', dtype=dtypes_spec)
    else: # .xlsx
        return pd.read_excel(io.BytesIO(file_content_bytes), dtype=dtypes_spec)

# --- Cached Helper Function: Clean and Prepare DataFrame ---
@st.cache_data
def clean_and_prepare_df(df_input, source_file_name, column_type_map):
    """Cleans and prepares a single DataFrame based on user-defined column types."""
    # Work on a copy to ensure cached input df (if any) is not mutated directly
    df = df_input.copy()
    
    # 1. Initial NA replacement (strings, blanks, "N/A")
    # This should happen before type-specific conversions
    df = df.applymap(
        lambda x: pd.NA if (
            pd.isna(x)
            or (isinstance(x, str) and (x.strip() == "" or x.strip().upper() == "N/A"))
        ) else x
    )

    # 2. Apply user-defined column types
    # Errors='coerce' will turn unparseable values into NaT (for dates) or NaN (for numerics)
    for col_name in df.columns:
        # Ensure col_name exists in column_type_map, otherwise, it might be an extra column
        # in a non-conformant file. The main processing loop should handle column matching.
        if col_name not in column_type_map:
            # This case should ideally be caught by column count/name checks before this function
            # For safety, treat as text or skip, or raise error
            df[col_name] = df[col_name].astype(str).replace({'nan': pd.NA, '<NA>': pd.NA})
            continue

        col_type = column_type_map.get(col_name, "Text") # Default to Text if somehow missed

        if col_type == "Date":
            df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
        elif col_type == "Numeric":
            df[col_name] = pd.to_numeric(df[col_name], errors='coerce')
        elif col_type == "Text":
            # Ensure it's string, replace 'nan' or '<NA>' string that might arise from initial load
            df[col_name] = df[col_name].astype(str).replace({'nan': pd.NA, '<NA>': pd.NA, '': pd.NA})
    
    # 3. Optimized removal of trailing empty rows (post-type conversion)
    # An "empty" cell is now pd.NA (or NaT for datetimes, NaN for numerics)
    
    if not df.empty:
        # A row is empty if all its cells are NA (pd.NA, NaT, NaN)
        all_empty_mask_rows = df.isna().all(axis=1)

        if all_empty_mask_rows.all(): # If all rows are empty
            df = df.iloc[0:0] # Return empty dataframe with original columns
        else:
            non_empty_idx = (~all_empty_mask_rows).to_numpy().nonzero()[0]
            if len(non_empty_idx) > 0:
                last_valid_idx = non_empty_idx[-1]
                df = df.iloc[:last_valid_idx + 1]
            # If non_empty_idx is empty, it means all rows were considered empty.
            # This case is covered by all_empty_mask_rows.all().
    
    df['Source_File'] = os.path.basename(source_file_name)

    # 4. Final type formatting for output (especially dates to string YYYY-MM-DD)
    for col_name in df.columns:
        if col_name == 'Source_File':
            continue
        
        # Use the type map again to ensure correct formatting
        col_type = column_type_map.get(col_name)
        if col_type == "Date" and col_name in df.columns:
            # Check if the column is actually datetime dtype after coercion
            if pd.api.types.is_datetime64_any_dtype(df[col_name]):
                # Format to string, NaT will become pd.NA after strftime, then handle pd.NA to None
                df[col_name] = df[col_name].dt.strftime('%Y-%m-%d')
            # After strftime, NaT becomes pd.NA (a string '<NA>' or similar if not handled carefully)
            # So, replace these specific string representations of NA if they occur
            df[col_name] = df[col_name].replace({'<NA>': pd.NA, 'NaT': pd.NA, '': pd.NA})


    return df

# --- Streamlit App ---
st.set_page_config(layout="wide")
st.title("Excel Consolidator to SQLite")

st.warning(
    "Important: All spreadsheet tables should have the "
    "same order, number, and naming of columns for the application to work correctly. "
    "The column names and order from the first valid Excel file will be used as the standard."
)

# Initialize session state variables
if 'column_names_standard' not in st.session_state:
    st.session_state.column_names_standard = [] # Columns from the first valid file
if 'user_column_types' not in st.session_state:
    st.session_state.user_column_types = {} # User's type choices for standard columns
if 'processed_df' not in st.session_state:
    st.session_state.processed_df = None
if 'excel_files_content' not in st.session_state:
    st.session_state.excel_files_content = [] # List of dicts {'name': file_name, 'content': file_bytes}
if 'current_file_names' not in st.session_state:
    st.session_state.current_file_names = []


uploaded_files = st.file_uploader(
    "Upload one or more Excel files (.xlsx, .xlsb)", 
    type=["xlsx", "xlsb"], 
    accept_multiple_files=True,
    key="file_uploader"
)

if uploaded_files:
    new_file_names = sorted([f.name for f in uploaded_files])
    # If new set of files uploaded, reset states
    if st.session_state.current_file_names != new_file_names:
        st.session_state.current_file_names = new_file_names
        st.session_state.excel_files_content = []
        st.session_state.column_names_standard = []
        st.session_state.user_column_types = {}
        st.session_state.processed_df = None
        
        first_excel_file_processed_for_cols = False
        
        for uploaded_file in uploaded_files:
            file_name = uploaded_file.name
            file_content_bytes = uploaded_file.getvalue() # Read file content
            st.session_state.excel_files_content.append({'name': file_name, 'content': file_content_bytes})
            
            if not first_excel_file_processed_for_cols:
                try:
                    temp_df_for_cols = None
                    if file_name.endswith('.xlsb'):
                        temp_df_for_cols = pd.read_excel(io.BytesIO(file_content_bytes), engine='pyxlsb')
                    else: # .xlsx
                        temp_df_for_cols = pd.read_excel(io.BytesIO(file_content_bytes))
                    
                    if not temp_df_for_cols.empty:
                        st.session_state.column_names_standard = temp_df_for_cols.columns.tolist()
                        temp_types = {}
                        for col in st.session_state.column_names_standard:
                            suggested_type = "Text"
                            col_lower = str(col).lower()
                            if any(kw in col_lower for kw in ['date', 'facture', 'échéance', 'dt', 'time']):
                                suggested_type = "Date"
                            elif any(kw in col_lower for kw in ['code', 'montant', 'prix', 'qté', 'taux', 'total', 'num', 'id', 'valeur', 'quantity', 'amount', 'sum', 'count']):
                                suggested_type = "Numeric"
                            temp_types[col] = suggested_type
                        st.session_state.user_column_types = temp_types
                        first_excel_file_processed_for_cols = True
                    else:
                        st.info(f"Uploaded file '{file_name}' was empty or unreadable for column detection. Will try next file if available.")
                except Exception as e:
                    st.warning(f"Could not read columns from '{file_name}' to set standard: {e}. Will try next file if available.")
        
        if not st.session_state.excel_files_content: # Should not happen if uploaded_files is not empty
            st.warning("No files were processed after upload.")
        elif not st.session_state.column_names_standard:
             st.error("Could not extract column headers from any of the uploaded Excel files. Ensure at least one file is valid and non-empty.")

elif not uploaded_files and st.session_state.current_file_names: 
    # If all files are cleared from uploader, reset everything
    st.session_state.current_file_names = []
    st.session_state.excel_files_content = []
    st.session_state.column_names_standard = []
    st.session_state.user_column_types = {}
    st.session_state.processed_df = None


if st.session_state.column_names_standard and st.session_state.excel_files_content:
    st.subheader("Step 1: Define Column Data Types")
    first_file_name_for_cols = ""
    if st.session_state.excel_files_content: # Find which file provided the columns
        for f_data in st.session_state.excel_files_content:
            try: # Quick check if this file could have been the one
                df_check = pd.read_excel(io.BytesIO(f_data['content']), engine='pyxlsb' if f_data['name'].endswith('.xlsb') else None, nrows=0)
                if df_check.columns.tolist() == st.session_state.column_names_standard:
                    first_file_name_for_cols = f_data['name']
                    break
            except:
                pass
    if not first_file_name_for_cols and st.session_state.excel_files_content: # Fallback if exact match not found quickly
        # Try to find the file that was used for column detection
        for f_data in st.session_state.excel_files_content:
            try:
                # A bit redundant, but confirms which file was likely used
                df_check = pd.read_excel(io.BytesIO(f_data['content']), 
                                         engine='pyxlsb' if f_data['name'].endswith('.xlsb') else None, 
                                         nrows=0)
                if df_check.columns.tolist() == st.session_state.column_names_standard:
                    first_file_name_for_cols = f_data['name']
                    break
            except:
                continue
        if not first_file_name_for_cols: # If still not found (e.g. first file was empty but another provided cols)
             first_file_name_for_cols = st.session_state.excel_files_content[0]['name']


    st.markdown(f"Column names are based on the first valid Excel file processed: **{first_file_name_for_cols}**. "
                "Please verify and adjust the data type for each column. These types will be enforced.")
    
    cols_per_row = 3
    grid_cols = st.columns(cols_per_row)
    
    # Create a temporary dictionary to hold changes from selectboxes for this render cycle
    current_selection_types = st.session_state.user_column_types.copy()

    for i, col_name in enumerate(st.session_state.column_names_standard):
        with grid_cols[i % cols_per_row]:
            default_options = ["Text", "Numeric", "Date"]
            # Get the current type from session state, default to "Text" if not found
            current_type_for_col = st.session_state.user_column_types.get(col_name, "Text")
            
            try:
                current_idx = default_options.index(current_type_for_col)
            except ValueError: 
                current_idx = default_options.index("Text") # Fallback if stored type is invalid
                st.session_state.user_column_types[col_name] = "Text" # Correct it in session state

            selected_type = st.selectbox(
                f"Type for '{col_name}'",
                options=default_options,
                index=current_idx,
                key=f"type_select_{col_name}" # Unique key for each selectbox
            )
            # Update the temporary dictionary with the selection
            current_selection_types[col_name] = selected_type
    
    # After all selectboxes are rendered, update the session state from the temporary dictionary
    # This is a common pattern to correctly handle widget state updates in Streamlit
    st.session_state.user_column_types = current_selection_types


    if st.button("Step 2: Process Files & Preview", key="process_button"):
        if not st.session_state.excel_files_content:
            st.error("No files to process. Please upload Excel files again.")
        else:
            all_dfs = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            total_files = len(st.session_state.excel_files_content)
            
            # Define dtypes for faster reading in the main processing loop.
            # All columns will be read as 'object' (strings), and clean_and_prepare_df will handle conversions.
            forced_object_dtypes = {col_name: 'object' for col_name in st.session_state.column_names_standard}
            
            for i, file_data in enumerate(st.session_state.excel_files_content):
                file_name_in_zip = file_data['name']
                file_content_bytes = file_data['content']
                status_text.text(f"Processing file {i+1}/{total_files}: {file_name_in_zip}...")
                
                try:
                    current_df_raw = None
                    # Load raw DataFrame using the cached function
                    current_df_raw = load_raw_df(file_content_bytes, file_name_in_zip, forced_object_dtypes)
                    
                    # Validate column structure against the standard
                    if list(current_df_raw.columns) != st.session_state.column_names_standard:
                        st.warning(f"Skipping '{file_name_in_zip}': Column names/order mismatch. "
                                   f"Expected: {st.session_state.column_names_standard}, "
                                   f"Found: {list(current_df_raw.columns)}. Ensure all files have the exact same structure as the first valid file.")
                        continue
                    
                    # Process the raw DataFrame (which itself might be a copy due to clean_and_prepare_df's internal .copy())
                    processed_single_df = clean_and_prepare_df(current_df_raw, file_name_in_zip, st.session_state.user_column_types)
                    all_dfs.append(processed_single_df)
                except Exception as e:
                    st.error(f"Error processing '{file_name_in_zip}': {e}")
                
                progress_bar.progress((i + 1) / total_files)
            
            status_text.text("Consolidating data...")
            if all_dfs:
                st.session_state.processed_df = pd.concat(all_dfs, ignore_index=True)
                
                # Final pass to ensure NAs are consistent (pd.NA) for display and SQLite prep
                for col in st.session_state.processed_df.columns:
                    # Convert various forms of string NA/empty to pd.NA if not already
                    if st.session_state.processed_df[col].dtype == 'object':
                         st.session_state.processed_df[col] = st.session_state.processed_df[col].replace({'nan': pd.NA, '<NA>': pd.NA, 'NaT': pd.NA, '': pd.NA})
                
                st.success("Files processed successfully!")
                status_text.empty()
                progress_bar.empty()
            else:
                st.error("No data could be processed from the files. Check individual file errors or warnings.")
                st.session_state.processed_df = None
                status_text.text("Processing failed or no valid data.")
                progress_bar.empty()


if st.session_state.processed_df is not None and not st.session_state.processed_df.empty:
    st.subheader("Preview of Consolidated Data (First 100 Rows)")
    # For display, make NAs more readable if needed, though Streamlit handles pd.NA well
    preview_df = st.session_state.processed_df.head(100).copy()
    # Streamlit usually displays pd.NA as blank or 'NA'. If specific string needed:
    # preview_df = preview_df.fillna("[NO DATA]") 
    st.dataframe(preview_df)

    st.subheader("Step 3: Download SQLite Database")
    
    try:
        # Create a copy for SQLite to avoid altering the displayed/session df
        df_for_sqlite = st.session_state.processed_df.copy()

        # Convert pd.NA/NaT to None for SQLite compatibility
        for col in df_for_sqlite.columns:
            if pd.api.types.is_object_dtype(df_for_sqlite[col]) or pd.api.types.is_string_dtype(df_for_sqlite[col]):
                 df_for_sqlite[col] = df_for_sqlite[col].where(pd.notnull(df_for_sqlite[col]), None)
            elif pd.api.types.is_datetime64_any_dtype(df_for_sqlite[col]): # Should be string already
                 df_for_sqlite[col] = df_for_sqlite[col].astype(object).where(pd.notnull(df_for_sqlite[col]), None)
            # Handle Pandas nullable dtypes (IntegerNA, BooleanNA)
            elif df_for_sqlite[col].dtype.name in ['Int64', 'Int32', 'Int16', 'Int8', 'UInt64', 'UInt32', 'UInt16', 'UInt8', 'boolean']:
                 if df_for_sqlite[col].hasnans:
                    df_for_sqlite[col] = df_for_sqlite[col].astype(object).where(pd.notnull(df_for_sqlite[col]), None)
        
        # Create an in-memory SQLite database
        mem_conn = sqlite3.connect(':memory:')
        df_for_sqlite.to_sql(name='DATA', con=mem_conn, if_exists='replace', index=False, chunksize=1000)
        
        # Create a temporary file to save the binary database
        with tempfile.NamedTemporaryFile(delete=False, suffix='.db') as tmpfile:
            # Connect to the temporary disk database
            disk_conn = sqlite3.connect(tmpfile.name)
            # Backup the in-memory database to the disk database
            mem_conn.backup(disk_conn)
            disk_conn.close()
            tmpfile_path = tmpfile.name # Save path for reading

        mem_conn.close()

        # Read the binary content of the temporary file
        with open(tmpfile_path, 'rb') as f:
            db_bytes = f.read()
        
        os.remove(tmpfile_path) # Clean up the temporary file

        st.download_button(
            label="Download consolidated_data.db",
            data=db_bytes,
            file_name="consolidated_data.db",
            mime="application/vnd.sqlite3"
        )
    except Exception as e:
        st.error(f"Error creating SQLite database for download: {e}")

elif st.session_state.processed_df is not None and st.session_state.processed_df.empty:
    st.info("Processing resulted in an empty dataset. Nothing to preview or download.")

st.markdown("---")
st.markdown("Streamlit Excel Consolidator")

