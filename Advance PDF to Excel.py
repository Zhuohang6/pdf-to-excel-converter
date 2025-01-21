import os
import camelot
import pandas as pd

def pdf_to_excel_advanced(pdf_folder, output_folder):

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)
    print(f"Output folder: {output_folder}")
    
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
            # Extract tables using Camelot
            tables = camelot.read_pdf(pdf_path, pages="all", flavor="stream")
            if tables:
                print(f"Found {len(tables)} table(s) in {pdf_file}.")
                # Combine all tables into one Excel file
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    for i, table in enumerate(tables):
                        df = table.df  # Convert table to DataFrame
                        df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False, header=False)
                print(f"Saved: {excel_path}")
            else:
                print(f"No tables found in {pdf_file}.")
        except Exception as e:
            print(f"Error processing {pdf_file}: {e}")


pdf_folder = r"C:\Users\PDF"  # Replace with your actual folder path
output_folder = r"C:\Users\Excels"  # Replace with your actual folder path
pdf_to_excel_advanced(pdf_folder, output_folder)
