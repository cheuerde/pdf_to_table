import pdfplumber
import pandas as pd
import os
from pathlib import Path
import csv
import sys
import platform
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from threading import Thread

class BankStatementProcessor(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Bank Statement PDF Processor")
        self.geometry("600x400")
        
        # Platform-specific setup
        self.configure_platform_specifics()
        
        # Create main frame
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        
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
        parent.columnconfigure(1, weight=1)
        
        ttk.Label(parent, text="Input Folder (PDF files):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_path_var = tk.StringVar()
        input_entry = ttk.Entry(parent, textvariable=self.input_path_var)
        input_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(parent, text="Browse", command=self.select_input_folder).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(parent, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_path_var = tk.StringVar()
        output_entry = ttk.Entry(parent, textvariable=self.output_path_var)
        output_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        ttk.Button(parent, text="Browse", command=self.select_output_folder).grid(row=1, column=2, padx=5, pady=5)

    def create_progress_area(self, parent):
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="5")
        progress_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        progress_frame.columnconfigure(0, weight=1)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', variable=self.progress_var)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)

        log_frame = ttk.Frame(progress_frame)
        log_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
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

    def extract_table_from_pdf(self, pdf_path):
        all_data = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        if table[0][0] == 'Date':
                            if not all_data:
                                all_data.extend(table[1:])
                            else:
                                all_data.extend(table[1:])
                        else:
                            all_data.extend(table)
            
            self.log_message(f"Extracted {len(all_data)} rows from {Path(pdf_path).name}")
            return all_data
        except Exception as e:
            self.log_message(f"Error processing {pdf_path}: {str(e)}")
            return []

    def process_dataframe(self, df):
        try:
            df = df.dropna(how='all')
            
            for col in ['Balance', 'Paid in', 'Paid out']:
                df[col] = df[col].replace({None: '0', '': '0', 'nan': '0'})
                df[col] = df[col].astype(str).str.replace('£', '').str.replace(',', '')
                df[col] = pd.to_numeric(df[col], errors='coerce')
            
            df['Date'] = pd.to_datetime(df['Date'], format='%d %b %Y', errors='coerce')
            df['Description'] = df['Description'].astype(str).str.replace('\n', ' ').str.strip()
            df = df.dropna(subset=['Date'])
            df = df.sort_values('Date')
            
            return df
        except Exception as e:
            self.log_message(f"Error processing dataframe: {str(e)}")
            return df

    def save_dataframe(self, df, filepath):
        df.to_csv(filepath, 
                  index=False,
                  quoting=csv.QUOTE_NONNUMERIC,
                  quotechar='"',
                  doublequote=True,
                  date_format='%Y-%m-%d',
                  float_format='%.2f')

    def process_files(self):
        try:
            Path(self.output_folder).mkdir(parents=True, exist_ok=True)
            
            pdf_files = list(Path(self.input_folder).glob("*.pdf"))
            if not pdf_files:
                self.log_message("No PDF files found in input folder")
                return

            all_dfs = []
            total_files = len(pdf_files)

            for i, pdf_file in enumerate(pdf_files, 1):
                self.log_message(f"Processing: {pdf_file.name}")
                table_data = self.extract_table_from_pdf(pdf_file)
                
                if table_data:
                    headers = ["Date", "Type", "Description", "Paid in", "Paid out", "Balance"]
                    df = pd.DataFrame(table_data, columns=headers)
                    df = self.process_dataframe(df)
                    df['Source_File'] = pdf_file.name
                    all_dfs.append(df)
                    
                    individual_csv = Path(self.output_folder) / f"{pdf_file.stem}.csv"
                    self.save_dataframe(df, individual_csv)
                    self.log_message(f"Saved: {individual_csv}")

                progress = (i / total_files) * 100
                self.progress_var.set(progress)

            if all_dfs:
                combined_df = pd.concat(all_dfs, ignore_index=True)
                duplicate_cols = combined_df.columns.tolist()
                duplicate_cols.remove('Source_File')
                combined_df = combined_df.drop_duplicates(subset=duplicate_cols)
                combined_df = combined_df.sort_values('Date')
                
                combined_csv = Path(self.output_folder) / "all_transactions_combined.csv"
                combined_excel = Path(self.output_folder) / "all_transactions_combined.xlsx"
                
                self.save_dataframe(combined_df, combined_csv)
                combined_df.to_excel(combined_excel, index=False)
                
                self.log_message(f"\nProcessing Summary:")
                self.log_message(f"Total PDFs processed: {len(all_dfs)}")
                self.log_message(f"Total transactions: {len(combined_df)}")
                self.log_message(f"Date range: {combined_df['Date'].min()} to {combined_df['Date'].max()}")
                self.log_message(f"Files saved in: {self.output_folder}")
                
                messagebox.showinfo("Success", "Processing completed successfully!")
            
        except Exception as e:
            self.log_message(f"Error: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        
        finally:
            self.process_btn.configure(state='normal')

def main():
    try:
        if platform.system().lower() == 'windows':
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except:
                pass
        
        app = BankStatementProcessor()
        
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
        messagebox.showerror("Fatal Error", str(e))
        sys.exit(1)

if __name__ == "__main__":
    main()