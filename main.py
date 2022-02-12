import os
import sys
import pdfplumber
import pandas as pd
from glob import glob
from tqdm import tqdm

# Specify input folder!
input_path = "D:/Programming/PDF2XLSX/Input"


if __name__ == '__main__':
    # Create output folder
    output_path = input_path.replace("Input", "Output")
    os.makedirs(output_path, exist_ok=True)

    # Get all PDFs in input path
    files = glob(f"{input_path}/*.pdf")

    # Iterate over PDFs, extract tables, load into dataframe and save to XLSX
    print("Extracting tables from PDFs and saving them to XLSXs...")
    for f in tqdm(files, file=sys.stdout):
        # Set output file name
        name = f.replace("Input", "Output").replace("pdf", "xlsx")

        # If XLSX already exists, skip
        if os.path.exists(name):
            continue

        # Extract table to dataframe
        with pdfplumber.open(f) as pdf:
            table = pdf.pages[0].extract_table()
            df = pd.DataFrame(table[1::], columns=table[0])

        # Save first two columns to XLSX
        df.to_excel(name, index=False, columns=["", "Your Score"])

    print("All done!")
