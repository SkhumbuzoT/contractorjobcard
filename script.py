import pandas as pd
from pdfrw import PdfReader, PdfWriter, PdfDict
import os

TEMPLATE = "NEW JOB CARD 2025 - BRUCE.pdf"
SPREADSHEET = "jobcard_data.xlsx"
OUTPUT_FOLDER = "output_cards/"

# Make sure output folder exists
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def fill_pdf(template_path, output_path, data_dict):
    pdf = PdfReader(template_path)
    annotations = pdf.pages[0]["/Annots"]

    for annot in annotations:
        if annot["/Subtype"] == "/Widget" and annot.get("/T"):
            key = annot["/T"][1:-1]  # remove parentheses

            if key in data_dict:
                annot.update(
                    PdfDict(V=str(data_dict[key]), AS=str(data_dict[key]))
                )

    PdfWriter().write(output_path, pdf)


def main():
    df = pd.read_excel(SPREADSHEET)

    for idx, row in df.iterrows():

        # -------------------------------------------
        # 🔥 MAP YOUR SPREADSHEET COLUMNS → PDF FIELDS
        # -------------------------------------------
        data = {
            "*Account number:": row["Contract Account"],
            "*Address": f"{row['House No']} {row['Street']}",
            "*Suburb": row["City"],

            "*Old Meter Number": row["Device"],
            "Meter Reading": row["Old Meter Readings"],

            "*New Meter Number": row["New Meter Number"],

            # Optional fields if needed:
            # "Comment": row["Status"],
            # "SDC Name": "ROODEPOORT",
            # "Installed by": "BRUCE MHLANGA",
            # "Company": "LIGHTUP ENTERPRISE"
        }

        # File name uses Contract Account to be unique
        output_file = f"{OUTPUT_FOLDER}JOB_CARD_{row['Contract Account']}.pdf"

        fill_pdf(TEMPLATE, output_file, data)

        print("Generated:", output_file)


if __name__ == "__main__":
    main()
