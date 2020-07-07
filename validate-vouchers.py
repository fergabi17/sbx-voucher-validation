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

vertical_code_patterns = {
    "orcb": "\n1\n0\n0\n0\n0\n0\n",
    "VOT_voucher_component_code": "V\n[0-9\n]+\n",
    "AW_voucher_component_code": "V\n-\n[0-9]{1,2}\n-\n[0-9\n]+-\n[A-Z\n]{4}"
}

horizontal_code_patterns = {
    "voucher_number": "\n[0-9]{12}\n",
    "voucher_number_format": "\n[0-9]{3} [0-9]{3} [0-9]{3} [0-9]{3}\n",
    "voucher_number_light": "\n[0-9]{9}\n",
    "voucher_key": "\n[0-9]{3}\n"
}


def get_voucher_code(code_name, code_pattern, code_position):
    code_list = re.findall(code_pattern, voucher_text)
    code = ""
    if len(code_list) == 1:
        code = code_list[0].replace("\n", "")
        if code_position == "vertical":
            # Vertical codes come inversed
            code = code[len(code)::-1]
    elif code_name == "orcb" and len(code_list) == 0:
        code_list = re.findall("\n\n1\n\n", voucher_text)
        if len(code_list) == 1:
            code = "1"
    else:
        code = "not found"
    all_voucher_codes[code_name] = code


def create_csv():
    time_stamp = re.findall("[0-9]+", str(datetime.timestamp(datetime.now())))
    csv_path = folder_path + "/vouchers_" + time_stamp[0] + ".csv"
    with open(csv_path, 'w', newline='') as file:
        fieldnames = []
        for key in all_voucher_codes:
            fieldnames.append(key)

        writer = csv.DictWriter(file, fieldnames=fieldnames)

        writer.writeheader()
        for voucher in voucher_list:
            writer.writerow(voucher)
        print("CSV created in: " + csv_path)


def print_voucher_codes():
    codes_string = "\n"
    for key in all_voucher_codes:
        codes_string += key + ": " + all_voucher_codes[key] + "\n"
    print(codes_string)


def validate_data(all_voucher_codes):
    val_list = []
    validation = True
    val_list.append(validate_orcb(all_voucher_codes["orcb"]))
    val_list.append(validate_vot_aw(all_voucher_codes["VOT_voucher_component_code"], all_voucher_codes["AW_voucher_component_code"]))
    val_list.append(validate_vn_format(all_voucher_codes["voucher_number"], all_voucher_codes["voucher_number_format"]))
    val_list.append(validate_vn_light(all_voucher_codes["voucher_number"], all_voucher_codes["voucher_number_light"]))
    val_list.append(validate_voucher_key(all_voucher_codes["voucher_number"], all_voucher_codes["voucher_key"]))
    val_list.append(validate_file_name(all_voucher_codes["file_name"], all_voucher_codes["VOT_voucher_component_code"]))
    if False in val_list:
        validation = False
    all_voucher_codes["validation"] = validation


def validate_orcb(orcb_code):
    if orcb_code == "000001" or orcb_code == "1":
        return True
    return False


def validate_vot_aw(vot_code, aw_code):
    vot_code = re.findall("[0-9]+", vot_code)[0]
    aw_code = re.findall("[0-9]{3,}", aw_code)[0]
    if vot_code == aw_code:
        return True
    return False


def validate_vn_format(vn_code, vn_format_code):
    vn_format_code = vn_format_code.replace(" ", "")
    if vn_format_code == vn_code:
        return True
    return False


def validate_vn_light(vn_code, vn_light_code):
    if vn_code[0:9] == vn_light_code:
        return True
    return False


def validate_voucher_key(vn_code, voucher_key_code):
    if vn_code[9:12] == voucher_key_code:
        return True
    return False


def validate_file_name(file_name, vot_code):
    file_name_pim = re.findall("[0-9]+", file_name)[1]
    vot_code_pim = vot_code.replace("V", "")
    if file_name_pim == vot_code_pim:
        return True
    return False


folder_path = easygui.diropenbox()
files_in_dir = os.listdir(folder_path)

print("\n• Those are the files in this directory: " + str(files_in_dir) + "\n")

# Main loop in pdf files found
if len(files_in_dir) >= 1:
    for pdf_file in files_in_dir:
        if pdf_file.endswith('.pdf'):
            all_voucher_codes = {
                "file_name": pdf_file
            }

            # Extracting text using PDF miner
            output_string = StringIO()
            with open(folder_path + "/" + pdf_file, 'rb') as in_file:
                parser = PDFParser(in_file)
                doc = PDFDocument(parser)
                rsrcmgr = PDFResourceManager()
                device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
                interpreter = PDFPageInterpreter(rsrcmgr, device)
                for page in PDFPage.create_pages(doc):
                    interpreter.process_page(page)

            voucher_text = output_string.getvalue()

            for key in vertical_code_patterns:
                get_voucher_code(key, vertical_code_patterns[key], "vertical")

            for key in horizontal_code_patterns:
                get_voucher_code(key, horizontal_code_patterns[key], "horizontal")

            validate_data(all_voucher_codes)
            voucher_list.append(all_voucher_codes.copy())

            #print(voucher_text)
            #print_voucher_codes()

#print(voucher_list)

create_csv()