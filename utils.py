import datetime
import os
import shutil
from typing import Any

import win32com.client as win32
from PyPDF2 import PdfWriter

# ___________________#
# ---| VARIABLES |---#
# -------------------#
excel_file_path: str = "./input_xlsx/masz.xlsx"
output_pdfs: str = "./output_pdfs/"
# sheet_name_1: str = "Tax Certificate Printing 2024"
sheet_name_1: str = "cert"
sheet_name_2: str = "SUM"
data_df_headers: list = [0, 1]
# DATA DF
data_password: tuple = ("last 4 digit of cnic", "Password")
data_Employee_Id_No: tuple = ("last 4 digit of cnic", "ID")
data_spell_no: tuple = ("last 4 digit of cnic", "SPELL NO")
data_address: tuple = ("last 4 digit of cnic", "ADDRESS")
data_name: tuple = ("last 4 digit of cnic", "NAME")
data_cnic: tuple = ("last 4 digit of cnic", "CNIC")
data_amount: tuple = ("last 4 digit of cnic", "AMOUNT ")
data_tax: tuple = ("last 4 digit of cnic", "tax")
# YY-MM (Multi-Index Column)
data_date_2023_7: tuple = ("2023-7", "Date")
data_amount_2023_7: tuple = ("2023-7", "AMOUNT")
data_tax_2023_7: tuple = ("2023-7", "TAX")
data_id_2023_7: tuple = ("2023-7", "ID")

data_date_2023_8: tuple = ("2023-8", "Date")
data_amount_2023_8: tuple = ("2023-8", "AMOUNT")
data_tax_2023_8: tuple = ("2023-8", "TAX")
data_id_2023_8: tuple = ("2023-8", "ID")

data_date_2023_9: tuple = ("2023-9", "Date")
data_amount_2023_9: tuple = ("2023-9", "AMOUNT")
data_tax_2023_9: tuple = ("2023-9", "TAX")
data_id_2023_9: tuple = ("2023-9", "ID")

data_date_2023_10: tuple = ("2023-10", "Date")
data_amount_2023_10: tuple = ("2023-10", "AMOUNT")
data_tax_2023_10: tuple = ("2023-10", "TAX")
data_id_2023_10: tuple = ("2023-10", "ID")

data_date_2023_11: tuple = ("2023-11", "Date")
data_amount_2023_11: tuple = ("2023-11", "AMOUNT")
data_tax_2023_11: tuple = ("2023-11", "TAX")
data_id_2023_11: tuple = ("2023-11", "ID")

data_date_2023_12: tuple = ("2023-12", "Date")
data_amount_2023_12: tuple = ("2023-12", "AMOUNT")
data_tax_2023_12: tuple = ("2023-12", "TAX")
data_id_2023_12: tuple = ("2023-12", "ID")

data_date_2024_1: tuple = ("2024-1", "Date")
data_amount_2024_1: tuple = ("2024-1", "AMOUNT")
data_tax_2024_1: tuple = ("2024-1", "TAX")
data_id_2024_1: tuple = ("2024-1", "ID")

data_date_2024_2: tuple = ("2024-2", "Date")
data_amount_2024_2: tuple = ("2024-2", "AMOUNT")
data_tax_2024_2: tuple = ("2024-2", "TAX")
data_id_2024_2: tuple = ("2024-2", "ID")

data_date_2024_3: tuple = ("2024-3", "Date")
data_amount_2024_3: tuple = ("2024-3", "AMOUNT")
data_tax_2024_3: tuple = ("2024-3", "TAX")
data_id_2024_3: tuple = ("2024-3", "ID")

data_date_2024_4: tuple = ("2024-4", "Date")
data_amount_2024_4: tuple = ("2024-4", "AMOUNT")
data_tax_2024_4: tuple = ("2024-4", "TAX")
data_id_2024_4: tuple = ("2024-4", "ID")

data_date_2024_5: tuple = ("2024-5", "Date")
data_amount_2024_5: tuple = ("2024-5", "AMOUNT")
data_tax_2024_5: tuple = ("2024-5", "TAX")
data_id_2024_5: tuple = ("2024-5", "ID")

data_date_2024_6: tuple = ("2024-6", "Date")
data_amount_2024_6: tuple = ("2024-6", "AMOUNT")
data_tax_2024_6: tuple = ("2024-6", "TAX")
data_id_2024_6: tuple = ("2024-6", "ID")

# CERT DF
# cert_Employee_Id_No: str = "Unnamed: 2"
cert_EMPLOYEE_ID_NO: str = "C3"
cert_TAX: str = "F5"
cert_SPELL_NO: str = "D7"
cert_NAME: str = "F8"
cert_ADDRESS: str = "F9"
cert_CNIC: str = "F13"
cert_AMOUNT: str = "F19"

cert_DATE_YY_7: str = "B28"
cert_DATE_TAX_YY_7: str = "H28"
cert_DATE_ID_YY_7: str = "J28"

cert_DATE_YY_8: str = "B29"
cert_DATE_TAX_YY_8: str = "H29"
cert_DATE_ID_YY_8: str = "J29"

cert_DATE_YY_9: str = "B30"
cert_DATE_TAX_YY_9: str = "H30"
cert_DATE_ID_YY_9: str = "J30"

cert_DATE_YY_10: str = "B31"
cert_DATE_TAX_YY_10: str = "H31"
cert_DATE_ID_YY_10: str = "J31"

cert_DATE_YY_11: str = "B32"
cert_DATE_TAX_YY_11: str = "H32"
cert_DATE_ID_YY_11: str = "J32"

cert_DATE_YY_12: str = "B33"
cert_DATE_TAX_YY_12: str = "H33"
cert_DATE_ID_YY_12: str = "J33"

cert_DATE_YY_1: str = "B34"
cert_DATE_TAX_YY_1: str = "H34"
cert_DATE_ID_YY_1: str = "J34"

cert_DATE_YY_2: str = "B35"
cert_DATE_TAX_YY_2: str = "H35"
cert_DATE_ID_YY_2: str = "J35"

cert_DATE_YY_3: str = "B36"
cert_DATE_TAX_YY_3: str = "H36"
cert_DATE_ID_YY_3: str = "J36"

cert_DATE_YY_4: str = "B37"
cert_DATE_TAX_YY_4: str = "H37"
cert_DATE_ID_YY_4: str = "J37"

cert_DATE_YY_5: str = "B38"
cert_DATE_TAX_YY_5: str = "H38"
cert_DATE_ID_YY_5: str = "J38"

cert_DATE_YY_6: str = "B39"
cert_DATE_TAX_YY_6: str = "H39"
cert_DATE_ID_YY_6: str = "J39"


# ___________________#
# ---| FUNCTIONS |---#
# -------------------#
# to format the datetime part of the MultiIndex
def format_datetime(index) -> str | Any:
    if isinstance(index, datetime.datetime):
        return f"{index.year}-{index.month}"
    return index


def ensure_clean_directory(directory_path) -> None:
    """
    Ensure that the specified directory exists and is empty. If the directory
    does not exist, it will be created. If it exists, all files and directories
    within it will be deleted.
    """
    if not os.path.exists(directory_path):
        os.mkdir(directory_path)  # Create the folder if it does not exist
    else:
        # If the folder exists, delete all files in the folder
        for filename in os.listdir(directory_path):
            file_path: str = os.path.join(directory_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)  # Remove the file or link
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)  # Remove the directory
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")


def convert_excel_to_pdf_with_password(
    excel_path, sheet_name, output_pdf_path, password
) -> None:
    excel_path: str = os.path.abspath(excel_path)  # Convert to absolute path
    output_pdf_path: str = os.path.abspath(output_pdf_path)

    # Convert Excel sheet to PDF
    excel: Any = win32.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        w_b: Any = excel.Workbooks.Open(excel_path)
    except Exception as e:
        print(f"Failed to open Excel file: {e}")
        excel.Quit()
        return

    # Access the specific sheet
    ws: Any = w_b.Sheets(sheet_name)

    # Save as PDF
    try:
        ws.ExportAsFixedFormat(0, output_pdf_path)
    except Exception as e:
        print(f"Failed to export as PDF: {e}")
        w_b.Close(False)
        excel.Quit()
        return

    # Close the workbook and quit Excel
    w_b.Close(False)
    excel.Quit()

    # Apply password protection to the PDF
    try:
        with open(output_pdf_path, "rb") as pdf_file:
            reader: Any = PdfWriter()
            reader.append(pdf_file)
            reader.encrypt(password)

            # Save the password-protected PDF, overwriting the original file
            with open(output_pdf_path, "wb") as protected_pdf_file:
                reader.write(protected_pdf_file)

        # print(f"PDF saved and password protected as: {output_pdf_path}")
    except Exception as e:
        print(f"Failed to protect PDF with password: {e}")
