from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from datetime import datetime

import re
import easygui
import os
import csv

voucher_list = []
vertical_codes = ["orcb", "VOT_voucher_component_code",
                  "AW_voucher_component_code"]

PIM_code_patterns = {
    "orcb": "\n1\n0\n0\n0\n0\n0\n",
    "VOT_voucher_component_code": "V\n[0-9\n]+\n",
    "AW_voucher_component_code": "V\n-\n[0-9]{1,2}\n-\n[0-9\n]+-\n[A-Z\n]{4}",
    "voucher_number": "\n[0-9]{12}\n",
    "voucher_number_format": "\n[0-9]{3} [0-9]{3} [0-9]{3} [0-9]{3}\n",
    "voucher_number_light": "\n[0-9]{9}\n",
    "voucher_key": "\n[0-9]{3}\n"
}

UK_code_patterns = {
    "orcb": "\n1\n0\n0\n0\n0\n0\n",
    "VOT_voucher_component_code": "V\n2\n1\n[0-9\n]{4}[A-Z]\n[0-9\n]{10}[A-Z\n]{6}",
    "AW_voucher_component_code": "V\n2\n1\n[0-9\n]{4}[A-Z]\n[0-9\n]{10}[A-Z\n]{6}",
    "voucher_number": "\n[0-9]{19}\n",
    "voucher_number_format": "\n[0-9]{19}\n",
    "voucher_number_light": "\n[0-9]{12}\n",
    "voucher_key": "\n[0-9]{6}\n"
}

PT_code_patterns = {
    "orcb": "\n1\n0\n0\n0\n0\n0\n",
    "VOT_voucher_component_code": "V\n2\n1\n[0-9\n]{4}[A-Z]\n[0-9\n]{6}[A-Z\n]{6}",
    "AW_voucher_component_code": "V\n2\n1\n[0-9\n]{4}[A-Z]\n[0-9\n]{6}[A-Z\n]{6}",
    "voucher_number": "\n[0-9]{9}\n",
    "voucher_number_format": "\nhttps://odissei.as/[A-z0-9]+\n",
    "voucher_number_light": "\n[0-9]{10}\n",
    "voucher_key": "\n[0-9]{4}\n"
}


def get_voucher_code(code_name, code_pattern):
    """
    Uses a regex pattern to find the code in the voucher_text string
    and returns it.
    """

    code_list = re.findall(code_pattern, voucher_text)
    if len(code_list) >= 1:
        code = code_list[0].replace("\n", "")
        if code_name in vertical_codes:
            # Vertical codes come inversed
            code = code[len(code)::-1]
        return code
    if code_name == "orcb" and code_pattern != "\n\n1\n\n":
        return get_voucher_code(code_name, "\n\n1\n\n")

    return "not found"


def create_csv():
    """
    Creates a csv from the voucher_list. This is a list of dictionaries containing the voucher codes.
    The csv is saved in the same location where the pdf files were found, with a timestamp.
    """

    time_stamp = re.findall("[0-9]+", str(datetime.timestamp(datetime.now())))
    csv_path = folder_path + "/vouchers_" + time_stamp[0] + ".csv"

    with open(csv_path, 'w', newline='') as file:
        fieldnames = []
        for key in voucher_data:
            fieldnames.append(key)

        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()

        for voucher in voucher_list:
            writer.writerow(voucher)


def get_voucher_data(pdf_file):
    """
    Returns a voucher object extracting information from the pdf name.
    """

    box_country = pdf_file.split("_")[1][:2]
    if str(box_country) != "UK" and str(box_country) != "PT":
        box_country = "PIM"

    return {
        "file_name": pdf_file,
        "country": box_country
    }


def extract_text(pdf_file):
    """
    Returns all text from a pdf file in a string using PDF miner.
    """

    output_string = StringIO()
    with open(folder_path + "/" + pdf_file, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(
            rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

    return output_string.getvalue()


def validate_data(voucher_data):
    """
    Checks if all codes were found per voucher.
    It also validates codes calling different functions.
    Only returns True if all the information is validated.
    """

    for code in voucher_data:
        if voucher_data[code] == "not found":
            return False

    if voucher_data["country"] == "PIM":
        val_list = [validate_vot_aw(), validate_vn_format(
        ), validate_vn_light(), validate_voucher_key(), validate_file_name()]
        if False in val_list:
            return False

    return True


"""
>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
All 'validate' functions bellow validate specificities for each type of code.
All of them are called in the validate_data() according to the country.
>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
"""

def validate_vot_aw():
    vot_code = re.findall(
        "[0-9]+", voucher_data["VOT_voucher_component_code"])[0]
    aw_code = re.findall(
        "[0-9]{3,}", voucher_data["AW_voucher_component_code"])[0]
    if vot_code == aw_code:
        return True
    return False


def validate_vn_format():
    vn_code = voucher_data["voucher_number"]
    vn_format_code = voucher_data["voucher_number_format"].replace(" ", "")
    if vn_format_code == vn_code:
        return True
    return False


def validate_vn_light():
    vn_code = voucher_data["voucher_number"]
    vn_light_code = voucher_data["voucher_number_light"]
    if vn_code[0:9] == vn_light_code:
        return True
    return False


def validate_voucher_key():
    vn_code = voucher_data["voucher_number"]
    voucher_key_code = voucher_data["voucher_key"]
    if vn_code[9:12] == voucher_key_code:
        return True
    return False


def validate_file_name():
    file_name_pim = re.findall("[0-9]+", voucher_data["file_name"])[1]
    vot_code_pim = voucher_data["VOT_voucher_component_code"].replace("V", "")
    if file_name_pim == vot_code_pim:
        return True
    return False


"""
>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
START THE PROCESS
>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
"""

folder_path = easygui.diropenbox()
files_in_dir = os.listdir(folder_path)

# Main loop in pdf files found
if len(files_in_dir) >= 1:

    for pdf_file in files_in_dir:
        if pdf_file.endswith('.pdf'):

            voucher_data = get_voucher_data(pdf_file)
            voucher_text = extract_text(pdf_file)
            code_patterns = eval(voucher_data["country"] + "_code_patterns")

            for key in code_patterns:
                voucher_data[key] = get_voucher_code(
                    key, code_patterns[key])

            voucher_data["validation"] = validate_data(voucher_data)

            # Make sure numbers will be displayed correctly in excel
            for key in voucher_data:
                voucher_data[key]  = f"\'{voucher_data[key]}"
            
            voucher_list.append(voucher_data.copy())

if len(voucher_list) >= 1:
    create_csv()

easygui.msgbox(f"Process Finished! Checked {len(voucher_list)} voucher(s).")
