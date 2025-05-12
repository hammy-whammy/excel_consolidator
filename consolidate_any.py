import os
import threading
import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
from sqlalchemy import create_engine
from multiprocessing import Pool, cpu_count
import time

def process_file(file):
    """Process a single .xlsx or .xlsb file appropriately."""
    try:
        if file.endswith('.xlsb'):
            df = pd.read_excel(file, engine='pyxlsb')
        elif file.endswith('.xlsx'):
            df = pd.read_excel(file)
        else:
            return None  # skip unknown file types

        # --- Optimized removal of trailing rows where all columns are blank, N/A, or zero (for numeric columns) ---
        # Replace empty strings, whitespace-only strings, and "N/A" (case-insensitive) with pd.NA
        df = df.applymap(
            lambda x: pd.NA if (
                pd.isna(x)
                or (isinstance(x, str) and (x.strip() == "" or x.strip().upper() == "N/A"))
            ) else x
        )

        obj_cols = df.select_dtypes(include=['object']).columns
        num_cols = df.select_dtypes(include=['number']).columns

        # Create boolean masks for "empty" cells
        obj_empty_mask = df[obj_cols].isna() if len(obj_cols) > 0 else None
        num_empty_mask = (df[num_cols].isna() | (df[num_cols] == 0)) if len(num_cols) > 0 else None

        # Combine masks to get "empty" rows
        if obj_empty_mask is not None and num_empty_mask is not None:
            all_empty_mask = obj_empty_mask.all(axis=1) & num_empty_mask.all(axis=1)
        elif obj_empty_mask is not None:
            all_empty_mask = obj_empty_mask.all(axis=1)
        elif num_empty_mask is not None:
            all_empty_mask = num_empty_mask.all(axis=1)
        else:
            all_empty_mask = pd.Series([True] * len(df), index=df.index)

        # Find the last non-empty row index (fast, vectorized)
        non_empty_idx = (~all_empty_mask).to_numpy().nonzero()[0]
        if len(non_empty_idx) > 0:
            last_valid_idx = non_empty_idx[-1]
            df = df.iloc[:last_valid_idx + 1]
        # -------------------------------------------------------------------------------------------

        file_name = os.path.basename(file)
        df['Source_File'] = file_name

        # Clean date columns to uniform YYYY-MM-DD format
        for date_col in ["Date\nFacture", "Date échéance"]:
            if date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.strftime('%Y-%m-%d')

        # Convert specified columns to numeric
        numeric_cols = [
            "Code Centre",
            "Code Client",
            "Qtés facturées",
            "Prix unitaire\nen euros HT",
            "Total\nen euros HT",
            "TAUX TVA",
            "Montant TVA\npar article"
        ]
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')

        # Uncomment the following lines if you want to create a Type column based on keywords
        # Create the Type column based on keywords found in the source file name
        # keywords = ["Colas", "Construction", "Telecom", "Immobilier", "TF1"]
        # found_type = None
        # for keyword in keywords:
        #     if keyword.lower() in file_name.lower():
        #         found_type = keyword
        #         break
        # df['Type'] = found_type

        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str)
        return df
    except Exception as e:
        print(f"Error processing {file}: {str(e)}")
        return None

class AnyExcelConsolidator:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Consolidator (.xlsx & .xlsb)")
        self.root.geometry("500x300")
        
        # GUI Elements
        Label(root, text="Select Folder:").pack(pady=10)
        
        self.folder_path = StringVar()
        Entry(root, textvariable=self.folder_path, width=50).pack(pady=5)
        Button(root, text="Browse", command=self.browse_folder).pack(pady=5)
        
        self.progress = Progressbar(root, orient=HORIZONTAL, length=300, mode='determinate')
        self.progress.pack(pady=20)
        
        self.status = Label(root, text="Ready")
        self.status.pack(pady=10)
        
        Button(root, text="Process Files", command=self.start_processing).pack(pady=10)
        
        self.num_workers = cpu_count()  # Use all available CPU cores
    
    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)
    
    def start_processing(self):
        if not self.folder_path.get():
            messagebox.showerror("Error", "Please select a folder first")
            return
        
        # Disable button during processing
        for widget in self.root.winfo_children():
            if isinstance(widget, Button) and widget.cget("text") == "Process Files":
                widget.config(state=DISABLED)
        
        # Start processing in a separate thread
        threading.Thread(target=self.process_files, daemon=True).start()
    
    def process_files(self):
        try:
            start_time = time.time()
            self.status.config(text="Searching for files...")
            self.root.update()
            
            # Recursive search for .xlsx and .xlsb files
            excel_files = [
                os.path.join(root_dir, file)
                for root_dir, _, files in os.walk(self.folder_path.get())
                for file in files if file.endswith('.xlsx') or file.endswith('.xlsb')
            ]
            
            if not excel_files:
                messagebox.showinfo("Info", "No .xlsx or .xlsb files found in the selected folder")
                return
            
            self.status.config(text=f"Found {len(excel_files)} files. Processing...")
            self.progress["maximum"] = len(excel_files)
            self.root.update()
            
            # Use multiprocessing to process files in parallel
            with Pool(self.num_workers) as pool:
                results = pool.map(process_file, excel_files)
            
            # Filter out None results (failed files)
            dfs = [df for df in results if df is not None]
            
            if not dfs:
                messagebox.showerror("Error", "No valid data found in any files")
                return
            
            final_df = pd.concat(dfs, ignore_index=True)
            
            # Replace "nan" strings with None to standardize missing values
            final_df = final_df.replace("nan", None)
            
            # Standardize null values: replace any remaining pd.NaN with None
            final_df = final_df.where(pd.notnull(final_df), None)

            # Convert all datetime columns to string to avoid SQLite errors
            for col in final_df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']).columns:
                final_df[col] = final_df[col].dt.strftime('%Y-%m-%d')

            # Also convert any Timestamp objects in object columns to string
            for col in final_df.select_dtypes(include=['object']).columns:
                if final_df[col].apply(lambda x: isinstance(x, pd.Timestamp)).any():
                    final_df[col] = final_df[col].apply(lambda x: x.strftime('%Y-%m-%d') if isinstance(x, pd.Timestamp) else x)

            # Save to SQLite
            db_path = os.path.join(self.folder_path.get(), "blanchissement_2020-24.db")
            engine = create_engine(f'sqlite:///{db_path}')
            
            self.status.config(text="Saving to database...")
            self.root.update()
            
            final_df.to_sql(
                name='DATA',
                con=engine,
                if_exists='replace',
                index=False
            )
            
            end_time = time.time()
            messagebox.showinfo("Success", f"Data successfully consolidated into:\n{db_path}\nProcessing time: {end_time - start_time:.2f} seconds")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        finally:
            self.status.config(text="Ready")
            self.progress["value"] = 0
            for widget in self.root.winfo_children():
                if isinstance(widget, Button) and widget.cget("text") == "Process Files":
                    widget.config(state=NORMAL)
            self.root.update()

if __name__ == "__main__":
    root = Tk()
    app = AnyExcelConsolidator(root)
    root.mainloop()
