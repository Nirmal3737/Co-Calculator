import camelot
import sys
import pandas as pd
def convert_pdf_to_excel(pdf_path, excel_path):
    # Extract tables from PDF
    tables = camelot.read_pdf(pdf_path, flavor='stream', pages='all')

    # Initialize an empty DataFrame
    df = pd.DataFrame()

    # Concatenate all tables into a single DataFrame
    for table in tables:
        df = pd.concat([df, table.df], ignore_index=True)


    # Write DataFrame to Excel file
    df.to_excel(excel_path, index=False)

    print(f"Excel file saved to {excel_path}")


# Example usage
pdf_file = sys.argv[1]
excel_file = sys.argv[2]
print(pdf_file,excel_file)
convert_pdf_to_excel(pdf_file, excel_file)
