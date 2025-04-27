import pdfplumber
import pandas as pd
import os
import csv
import sys
import platform
import traceback
import re
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread

# Custom Path implementation with fallback
try:
    from pathlib import Path
except ImportError:
    class Path:
        def __init__(self, path):
            self.path = str(path)
        
        def __truediv__(self, other):
            return Path(os.path.join(self.path, str(other)))
        
        def mkdir(self, parents=False, exist_ok=False):
            os.makedirs(self.path, exist_ok=exist_ok)
        
        def glob(self, pattern):
            import glob
            return [Path(p) for p in glob.glob(os.path.join(self.path, pattern))]
        
        @property
        def name(self):
            return os.path.basename(self.path)
        
        @property
        def stem(self):
            return os.path.splitext(self.name)[0]
        
        def __str__(self):
            return self.path

# CLI version of processor that doesn't require Tkinter
class CLIProcessor:
    def __init__(self, input_folder=None, output_folder=None):
        self.input_folder = input_folder
        self.output_folder = output_folder
    
    def log_message(self, message):
        print(message)
    
    def extract_table_from_pdf(self, pdf_path):
        all_data = []
        try:
            with pdfplumber.open(str(pdf_path)) as pdf:  # Convert Path to string
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        if table[0][0] == 'Date':
                            if not all_data:
                                all_data.extend(table)  # Include header row for first table
                            else:
                                all_data.extend(table[1:])  # Skip header for subsequent tables
                        else:
                            all_data.extend(table)
            
            # Ensure we have headers if data exists
            if all_data and all_data[0][0] == 'Date':
                headers = all_data.pop(0)  # Remove header row
                self.log_message(f"Extracted {len(all_data)} rows from {Path(pdf_path).name}")
                return headers, all_data
            elif all_data:
                self.log_message(f"Extracted {len(all_data)} rows from {Path(pdf_path).name}")
                return ["Date", "Type", "Description", "Paid in", "Paid out", "Balance"], all_data
            else:
                self.log_message(f"No data found in {Path(pdf_path).name}")
                return None, []
        except Exception as e:
            self.log_message(f"Error processing {pdf_path}: {str(e)}")
            return None, []
    
    def extract_account_number(self, filename):
        """Extract account number from filename pattern like 'Transactions--601730-01606158--16-12-2023-10-12-2024.pdf'"""
        try:
            # Look for pattern like 601730-01606158
            match = re.search(r'--(\d+)-(\d+)--', filename)
            if match:
                sort_code = match.group(1)
                account_number = match.group(2)
                return f"{sort_code}-{account_number}"
            else:
                return "Unknown"
        except Exception:
            return "Unknown"
    
    def parse_description(self, row):
        """Parse the description field for DPC and POS transactions"""
        desc_type = row.get('Type', '')
        description = row.get('Description', '')
        
        # Initialize all parsed fields as empty
        parsed = {
            'DPC1': '', 'DPC2': '', 'DPC3': '', 'DPC4': '', 'DPC5': '',
            'POS1': '', 'POS2': '', 'POS3': '', 'POS4': ''
        }
        
        if desc_type == 'DPC' and description:
            # DPC format typically has 5 parts separated by commas
            parts = [part.strip() for part in description.split(',')]
            for i, part in enumerate(parts[:5]):
                parsed[f'DPC{i+1}'] = part
                
        elif desc_type == 'POS' and description:
            # POS format typically has 4 parts separated by commas
            parts = [part.strip() for part in description.split(',')]
            for i, part in enumerate(parts[:4]):
                parsed[f'POS{i+1}'] = part
                
        return parsed

    def save_dataframe(self, df, filepath):
        df.to_csv(str(filepath),  # Convert Path to string
                  index=False,
                  quoting=csv.QUOTE_NONNUMERIC,
                  quotechar='"',
                  doublequote=True,
                  date_format='%Y-%m-%d',
                  float_format='%.2f')
    
    def create_balance_validation_file(self, df, filepath):
        """Create a file that helps validate balance calculations"""
        try:
            # Sort by account number and date
            df = df.sort_values(['Account_Number', 'Date']).copy()
            
            # Group by account number and calculate within each group
            grouped = df.groupby('Account_Number')
            result_dfs = []
            
            for account, group_df in grouped:
                group_df = group_df.sort_values('Date').copy()
                group_df['Next_Balance'] = group_df['Balance'].shift(-1)
                group_df['Calc_Next_Balance'] = group_df['Balance'] + group_df['Paid in'] - group_df['Paid out']
                group_df['Balance_Diff'] = group_df['Next_Balance'] - group_df['Calc_Next_Balance']
                group_df['Has_Discrepancy'] = abs(group_df['Balance_Diff']) > 0.01
                result_dfs.append(group_df)
            
            if result_dfs:
                final_df = pd.concat(result_dfs)
                final_df.to_csv(str(filepath), index=False)
                
                # Report statistics
                discrepancies = final_df['Has_Discrepancy'].sum()
                if discrepancies > 0:
                    self.log_message(f"Found {discrepancies} potential balance discrepancies")
                    self.log_message(f"Check {filepath} for details")
                else:
                    self.log_message("No balance discrepancies found")
            else:
                self.log_message("No data to validate")
                
        except Exception as e:
            self.log_message(f"Error creating balance validation file: {str(e)}")
    
    def process_files(self):
        try:
            Path(self.output_folder).mkdir(parents=True, exist_ok=True)
            
            pdf_files = list(Path(self.input_folder).glob("*.pdf"))
            if not pdf_files:
                self.log_message("No PDF files found in input folder")
                return False

            all_dfs = []
            total_files = len(pdf_files)

            for i, pdf_file in enumerate(pdf_files, 1):
                self.log_message(f"Processing ({i}/{total_files}): {pdf_file.name}")
                
                # Extract account number from filename
                account_number = self.extract_account_number(pdf_file.name)
                
                headers, table_data = self.extract_table_from_pdf(pdf_file)
                
                if table_data:
                    # Create DataFrame from the extracted data
                    df = pd.DataFrame(table_data, columns=headers)
                    
                    # Basic data cleaning
                    df = df.dropna(how='all')
                    
                    # Clean numeric values for display
                    for col in ['Balance', 'Paid in', 'Paid out']:
                        df[col] = df[col].replace({None: '0', '': '0', 'nan': '0'})
                        df[col] = df[col].astype(str).str.replace('£', '').str.replace(',', '')
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                    
                    # Process dates
                    df['Date'] = pd.to_datetime(df['Date'], format='%d %b %Y', errors='coerce')
                    df['Description'] = df['Description'].astype(str).str.replace('\n', ' ').str.strip()
                    df = df.dropna(subset=['Date'])
                    
                    # Add source file information, account number, and original order
                    df['File_Path'] = str(pdf_file)
                    df['Source_File'] = pdf_file.name
                    df['Account_Number'] = account_number
                    df['Original_Order'] = range(len(df))
                    
                    # Parse descriptions for DPC and POS types
                    parsed_desc = df.apply(self.parse_description, axis=1)
                    parsed_df = pd.DataFrame(list(parsed_desc))
                    
                    # Concatenate the parsed descriptions to the main dataframe
                    df = pd.concat([df, parsed_df], axis=1)
                    
                    all_dfs.append(df)
                    
                    # Save individual file
                    individual_csv = Path(self.output_folder) / f"{pdf_file.stem}.csv"
                    # Keep original PDF ordering for individual files
                    df_to_save = df.sort_values('Original_Order')
                    self.save_dataframe(df_to_save, individual_csv)
                    self.log_message(f"Saved: {individual_csv}")

                self.log_message(f"Progress: {i}/{total_files} files processed ({int(i/total_files*100)}%)")

            if all_dfs:
                combined_df = pd.concat(all_dfs, ignore_index=True)
                
                # Create a comprehensive duplicate detection key
                duplicate_cols = ['Date', 'Account_Number', 'Description', 'Type', 'Paid in', 'Paid out', 'Balance']
                combined_df = combined_df.drop_duplicates(subset=duplicate_cols)
                
                # Create transaction ID for more reliable identification
                combined_df['Transaction_ID'] = combined_df.apply(
                    lambda x: f"{x['Account_Number']}_{x['Date'].strftime('%Y-%m-%d')}_{x['Description']}_{x['Paid in']}_{x['Paid out']}",
                    axis=1
                )
                
                # Sort by account number and date
                combined_df = combined_df.sort_values(['Account_Number', 'Date'])
                
                # Remove helper columns before saving (but keep Original_Order for reference)
                combined_df = combined_df.drop(['Transaction_ID'], axis=1)
                
                # Reorder columns to put Account_Number and File_Path first
                cols = combined_df.columns.tolist()
                # Remove these columns from their current positions
                for col in ['Account_Number', 'File_Path', 'Source_File']:
                    if col in cols:
                        cols.remove(col)
                # Add them at the beginning
                cols = ['Account_Number', 'File_Path', 'Source_File'] + cols
                combined_df = combined_df[cols]
                
                combined_csv = Path(self.output_folder) / "all_transactions_combined.csv"
                combined_excel = Path(self.output_folder) / "all_transactions_combined.xlsx"
                
                self.save_dataframe(combined_df, combined_csv)
                combined_df.to_excel(str(combined_excel), index=False)
                
                self.log_message(f"\nProcessing Summary:")
                self.log_message(f"Total PDFs processed: {len(pdf_files)}")
                self.log_message(f"Total transactions: {len(combined_df)}")
                self.log_message(f"Date range: {combined_df['Date'].min()} to {combined_df['Date'].max()}")
                self.log_message(f"Accounts found: {', '.join(combined_df['Account_Number'].unique())}")
                self.log_message(f"Files saved in: {self.output_folder}")
                
                # Also save a file with balance validation information for analysis
                self.create_balance_validation_file(combined_df, Path(self.output_folder) / "balance_validation.csv")
                
                return True
            
        except Exception as e:
            self.log_message(f"Error: {str(e)}")
            traceback_info = traceback.format_exc()
            self.log_message(traceback_info)
            return False
        
        return False


class PDFProcessor(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PDF Processor")
        self.geometry("800x600")
        
        # Platform-specific setup
        self.configure_platform_specifics()
        
        # Create main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Folder selection frames
        self.create_folder_selection(main_frame)
        
        # Progress area
        self.create_progress_area(main_frame)
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="Process Files", command=self.start_processing)
        self.process_btn.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Status label
        self.status_var = tk.StringVar(value="Ready to process files")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=5, column=0, columnspan=3)

        self.input_folder = None
        self.output_folder = None

    def configure_platform_specifics(self):
        """Configure platform-specific settings"""
        system = platform.system().lower()
        
        if system == 'linux':
            try:
                self.tk.call('tk', 'scaling', 1.0)
                style = ttk.Style()
                try:
                    style.theme_use('clam')
                except:
                    pass
            except Exception as e:
                print(f"Warning: Could not apply Linux-specific settings: {e}")
                
        elif system == 'darwin':
            try:
                self.tk.call('tk', 'scaling', 2.0)
            except Exception as e:
                print(f"Warning: Could not apply macOS-specific settings: {e}")

    def create_folder_selection(self, parent):
        # Input folder selection
        ttk.Label(parent, text="Input Folder (PDF files):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_path_var = tk.StringVar()
        input_entry = ttk.Entry(parent, textvariable=self.input_path_var)
        input_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(parent, text="Browse", command=self.select_input_folder).grid(row=0, column=2, padx=5, pady=5)

        # Output folder selection
        ttk.Label(parent, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_path_var = tk.StringVar()
        output_entry = ttk.Entry(parent, textvariable=self.output_path_var)
        output_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(parent, text="Browse", command=self.select_output_folder).grid(row=1, column=2, padx=5, pady=5)

    def create_progress_area(self, parent):
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="5")
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        progress_frame.columnconfigure(0, weight=1)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', variable=self.progress_var)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)

        # Log area
        log_frame = ttk.Frame(progress_frame)
        log_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=15, width=70, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def select_input_folder(self):
        folder = filedialog.askdirectory(title="Select Input Folder")
        if folder:
            self.input_folder = folder
            self.input_path_var.set(folder)

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            self.output_path_var.set(folder)

    def log_message(self, message):
        print(message)
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    def start_processing(self):
        if not self.input_folder or not self.output_folder:
            messagebox.showerror("Error", "Please select both input and output folders")
            return

        self.process_btn.configure(state='disabled')
        self.progress_var.set(0)
        self.log_text.delete(1.0, tk.END)
        
        Thread(target=self.process_files, daemon=True).start()

    def process_files(self):
        # Create a CLI processor instance and delegate to it
        cli_processor = CLIProcessor(self.input_folder, self.output_folder)
        
        # Override the log message method to also log to the GUI
        original_log = cli_processor.log_message
        def gui_log(message):
            original_log(message)
            self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
            self.log_text.see(tk.END)
            self.update_idletasks()
            
            # Update progress bar for processing steps
            if "Progress:" in message:
                try:
                    progress_parts = message.split("(")[1].split("%")[0]
                    self.progress_var.set(float(progress_parts))
                except:
                    pass
        
        cli_processor.log_message = gui_log
        
        # Process the files
        success = cli_processor.process_files()
        
        # Update the GUI based on the result
        if success:
            messagebox.showinfo("Success", "Processing completed successfully!")
        
        self.process_btn.configure(state='normal')


def main():
    try:
        # Check if command-line arguments are provided
        if len(sys.argv) > 1:
            # Command-line mode
            if len(sys.argv) != 3:
                print("Usage: python extract.py <input_folder> <output_folder>")
                input("Press Enter to exit...")
                sys.exit(1)
            
            input_folder = sys.argv[1]
            output_folder = sys.argv[2]
            
            print(f"Processing PDFs from {input_folder} to {output_folder}")
            
            # Use the CLI processor instead of GUI
            processor = CLIProcessor(input_folder, output_folder)
            success = processor.process_files()
            
            if success:
                print("Processing complete!")
                input("Press Enter to exit...")
                sys.exit(0)
            else:
                print("Processing failed!")
                input("Press Enter to exit...")
                sys.exit(1)
        
        # GUI mode
        if platform.system().lower() == 'windows':
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except:
                pass
        
        app = PDFProcessor()
        
        # Center the window on screen
        app.update_idletasks()
        width = app.winfo_width()
        height = app.winfo_height()
        x = (app.winfo_screenwidth() // 2) - (width // 2)
        y = (app.winfo_screenheight() // 2) - (height // 2)
        app.geometry(f'{width}x{height}+{x}+{y}')
        
        app.mainloop()
    except Exception as e:
        print(f"Fatal error: {e}")
        traceback_info = traceback.format_exc()
        print(traceback_info)
        
        if not len(sys.argv) > 1:  # Only show messagebox in GUI mode
            try:
                messagebox.showerror("Fatal Error", f"{str(e)}\n\nSee console for details.")
            except:
                pass
        
        print("\nPress Enter to exit...")
        input()  # Wait for user input before exiting
        sys.exit(1)

if __name__ == "__main__":
    main()