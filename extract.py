import pdfplumber
import pandas as pd
import os
from pathlib import Path
import csv
import sys
import platform
from tkinter import filedialog, Tk
from datetime import datetime

def get_platform():
    return platform.system().lower()

def select_folder(title):
    root = Tk()
    root.withdraw()  # Hide the main window
    
    # Platform-specific adjustments
    if get_platform() == 'linux':
        try:
            root.attributes('-type', 'dialog')  # Linux-specific
        except:
            pass  # Ignore if the attribute isn't available
    
    folder = filedialog.askdirectory(title=title)
    return folder if folder else None

def extract_table_from_pdf(pdf_path):
    all_data = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # Process each page
            for page in pdf.pages:
                # Extract table from the current page
                table = page.extract_table()
                if table:
                    # Check if this is a header row
                    if table[0][0] == 'Date':
                        # Skip header row unless it's the first page and we have no data
                        if not all_data:
                            all_data.extend(table[1:])
                        else:
                            all_data.extend(table[1:])
                    else:
                        all_data.extend(table)
            
            print(f"Extracted {len(all_data)} rows from {Path(pdf_path).name}")
            return all_data
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return []

def process_dataframe(df):
    """Clean and process the dataframe"""
    try:
        # Remove any completely empty rows
        df = df.dropna(how='all')
        
        # Clean up the data
        for col in ['Balance', 'Paid in', 'Paid out']:
            # Replace empty strings and None with '0'
            df[col] = df[col].replace({None: '0', '': '0', 'nan': '0'})
            df[col] = df[col].astype(str).str.replace('£', '').str.replace(',', '')
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Convert date to datetime
        df['Date'] = pd.to_datetime(df['Date'], format='%d %b %Y', errors='coerce')
        
        # Clean up Description field - replace newlines with spaces
        df['Description'] = df['Description'].astype(str).str.replace('\n', ' ').str.strip()
        
        # Remove rows where Date is NaT (Not a Time)
        df = df.dropna(subset=['Date'])
        
        # Sort by date
        df = df.sort_values('Date')
        
        return df
    except Exception as e:
        print(f"Error processing dataframe: {str(e)}")
        return df

def save_dataframe(df, filepath):
    """Save dataframe with proper quoting"""
    df.to_csv(filepath, 
              index=False,
              quoting=csv.QUOTE_NONNUMERIC,
              quotechar='"',
              doublequote=True,
              date_format='%Y-%m-%d',
              float_format='%.2f')

def print_with_timestamp(message):
    """Print message with timestamp"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{timestamp}] {message}")

def main():
    try:
        # Platform-specific setup
        if get_platform() == 'windows':
            os.system('color')  # Enable colors in Windows command prompt

        print_with_timestamp("Bank Statement PDF Processor")
        print_with_timestamp("Please select the folder containing your PDF files...")
        
        pdf_folder = select_folder("Select PDF Input Folder")
        if not pdf_folder:
            print_with_timestamp("No input folder selected. Exiting...")
            return

        print_with_timestamp("Please select where to save the output files...")
        output_folder = select_folder("Select Output Folder")
        if not output_folder:
            print_with_timestamp("No output folder selected. Exiting...")
            return

        # Create output folder if it doesn't exist
        Path(output_folder).mkdir(parents=True, exist_ok=True)
        
        # List to store all dataframes
        all_dfs = []
        
        # Process each PDF file in the folder
        pdf_files = list(Path(pdf_folder).glob("*.pdf"))
        print_with_timestamp(f"Found {len(pdf_files)} PDF files to process")
        
        for pdf_file in pdf_files:
            print_with_timestamp(f"\nProcessing: {pdf_file}")
            
            # Extract table data from PDF
            table_data = extract_table_from_pdf(pdf_file)
            
            if table_data:
                # Create a DataFrame from the extracted data
                headers = ["Date", "Type", "Description", "Paid in", "Paid out", "Balance"]
                df = pd.DataFrame(table_data, columns=headers)
                
                # Process the dataframe
                df = process_dataframe(df)
                
                # Add source file information
                df['Source_File'] = pdf_file.name
                
                # Add to list of dataframes
                all_dfs.append(df)
                
                # Save individual file data
                individual_csv = Path(output_folder) / f"{pdf_file.stem}.csv"
                save_dataframe(df, individual_csv)
                print_with_timestamp(f"Saved individual file: {individual_csv}")
                print_with_timestamp(f"Extracted {len(df)} valid transactions")

        if all_dfs:
            # Combine all dataframes
            combined_df = pd.concat(all_dfs, ignore_index=True)
            
            # Remove duplicates based on all columns except Source_File
            duplicate_cols = combined_df.columns.tolist()
            duplicate_cols.remove('Source_File')
            combined_df = combined_df.drop_duplicates(subset=duplicate_cols)
            
            # Sort by date
            combined_df = combined_df.sort_values('Date')
            
            # Save the combined DataFrame
            combined_csv = Path(output_folder) / "all_transactions_combined.csv"
            combined_excel = Path(output_folder) / "all_transactions_combined.xlsx"
            
            # Save with proper quoting
            save_dataframe(combined_df, combined_csv)
            
            # Save Excel version
            combined_df.to_excel(combined_excel, index=False)
            
            # Print summary
            print_with_timestamp("\nProcessing Summary:")
            print_with_timestamp(f"Total number of PDFs processed: {len(all_dfs)}")
            print_with_timestamp(f"Total number of transactions: {len(combined_df)}")
            print_with_timestamp(f"Date range: {combined_df['Date'].min()} to {combined_df['Date'].max()}")
            print_with_timestamp(f"\nFiles saved in: {output_folder}")
            print_with_timestamp(f"Combined CSV: {combined_csv}")
            print_with_timestamp(f"Combined Excel: {combined_excel}")
            
            # Print transaction counts by source file
            print_with_timestamp("\nTransactions per file:")
            print(combined_df['Source_File'].value_counts())

        else:
            print_with_timestamp("No data was extracted from the PDF files")

    except Exception as e:
        print_with_timestamp(f"An error occurred: {str(e)}")
        raise

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print_with_timestamp("\nProcess interrupted by user")
    except Exception as e:
        print_with_timestamp(f"Fatal error: {str(e)}")
    finally:
        print_with_timestamp("\nProcess completed")
        input("Press Enter to exit...")
