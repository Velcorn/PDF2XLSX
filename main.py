import os
import sys
import pdfplumber
import pandas as pd
from glob import glob
from tqdm import tqdm

# Specify input_path and file_name!
# NOTE: r"{path}" - without curly brackets - is used for literal paths on Windows (backslash compatibility)
input_path = "D:/Programming/PDF2XLSX/Input"
file_name = "Output/Test.xlsx"


def pdf2xlsx():
    # Get all PDFs in input path
    files = glob(f"{input_path}/*.pdf")

    # Iterate over PDFs, extract tables and save desired column(s) into dataframe
    print("Extracting tables to dataframe...")
    df = pd.DataFrame()
    for i, f in enumerate(tqdm(files, file=sys.stdout)):
        # Extract table to dataframe
        with pdfplumber.open(f) as pdf:
            table = pdf.pages[0].extract_table()
            # Temporarily save entire table
            temp = pd.DataFrame(table[1::], columns=table[0])
            # If first PDF, add first and second column, else only second with file name as header
            name = f.split("/")[-1].split("\\")[-1].replace(".pdf", "")
            if i == 0:
                df[""] = temp[""]
                df[name] = temp["Your Score"].astype(float)
            else:
                df[name] = temp["Your Score"].astype(float)

    # Save dataframe to XLSX
    print("Saving dataframe to XLSX...")
    try:
        df.to_excel(file_name, index=False)
    # Catch error if file is open in another process
    except IOError as e:
        return f"{file_name} is used by another process, please close it first!"

    return "All done!"


if __name__ == '__main__':
    print(pdf2xlsx())
