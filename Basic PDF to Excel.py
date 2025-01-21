import os
import pdfplumber
import pandas as pd

def pdf_to_excel_batch(pdf_folder, output_folder):

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)
    print(f"Output folder: {output_folder}")
    
    # Debug the input folder path
    print(f"PDF folder exists: {os.path.exists(pdf_folder)}")
    print(f"Files in PDF folder: {os.listdir(pdf_folder)}")
    
    # List all PDF files in the folder (case-insensitive match)
    pdf_files = [file for file in os.listdir(pdf_folder) if file.lower().endswith('.pdf')]
    print(f"PDF files found: {pdf_files}")
    
    if not pdf_files:
        print("No PDF files found in the folder.")
        return
    
    for pdf_file in pdf_files:
        print(f"Processing {pdf_file}...")
        pdf_path = os.path.join(pdf_folder, pdf_file)
        excel_path = os.path.join(output_folder, pdf_file.replace('.PDF', '.xlsx').replace('.pdf', '.xlsx'))
        
        try:
            # Read the PDF and extract tables
            with pdfplumber.open(pdf_path) as pdf:
                all_dataframes = []
                
                for page in pdf.pages:
                    table = page.extract_table()
                    print(f"Page {page.page_number} - Table: {table}")
                    if table:
                        df = pd.DataFrame(table)
                        all_dataframes.append(df)
                
                if all_dataframes:
                    # Combine all tables into one Excel sheet
                    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                        for i, df in enumerate(all_dataframes):
                            df.to_excel(writer, sheet_name=f'Page_{i+1}', index=False, header=False)
                    print(f"Saved: {excel_path}")
                else:
                    print(f"No tables found in {pdf_file}")
        except Exception as e:
            print(f"Error processing {pdf_file}: {e}")


pdf_folder = r"C:\Users\PDF"  # Replace with your actual folder path
output_folder = r"C:\Users/Excel"  # Replace with your actual folder path
pdf_to_excel_batch(pdf_folder, output_folder)
