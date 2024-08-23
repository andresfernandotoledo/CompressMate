# -*- coding: utf-8 -*-
"""
Contains framework to upen Zip a zip file which contains multiple pdf files
representing invoices. The invoices are parsed into Invoice and Entry (invoice
entries) classes. These are converted to XLSX format.
"""

import os
import shutil
import zipfile
from subprocess import run
from PyPDF2 import PdfFileReader
from openpyxl import Workbook
from .logger import StatLogger
from .config import config
from .invoice import EntryTuple, invo_parser
from .utility import list2row
from .order_detail_xlsx_parse import GetOrderDetail, read_xlsx, write_xlsx

#[TODO] Put this to a manager class???
def pdf2rawtxt(pdfile, logger):
    """
    Read out the given pdf file to Invoice and Entry classes to parse it. Utilize
    PyPFD2 PdfFileReader. Go through every page of the pdf. When a new invoice
    entry was found by the Entry.parse_line it is appended to the Invoice.entries

    :param str pdfile: file path of the pdf to process
    :param logger: :class:`StatLogger`, collect statistical data about parsing

    :return: The invoice entry filled up with the information from pdf file
    :rtype: :class:`Invoice`
    """
    with open(pdfile, 'rb') as filedesc:
        return invo_parser(PdfFileReader(filedesc), logger)

def _init_clean_up(tmp_dir='tmp'):
    """
    Create tmp directory, delete it first if it already exists, if possible

    :param str tmp_dir: this is the directory to work during the execution
    """
    try:
        shutil.rmtree(tmp_dir)
    except FileNotFoundError:
        #Do not print any silly message as this is the expected behaviour
        pass
    finally:
        os.mkdir(tmp_dir)

def extract_zip(src_name, directory):
    """
    Extract the pdf files from the zip, there is no sanity check for the arguments
    :param str src_name: Path to a zip file to extract
    :param str dir: Path to the target directory to extract the zip file
    """
    with zipfile.ZipFile(src_name) as myzip:
        myzip.extractall(directory)

def get_pdf_files(directory, extension='.pdf'):
    """
    Walks through the given **dir** and collects every files with **extension**

    :param str dir: the root directory to start the walk
    :param str extension: '.pdf' by default, if the file has this extension it is selected

    :return: list of pdf file path
    :rtype: list of str
    """
    pdf_list = []
    for dir_path, _dummy, file_list in os.walk(directory):
        for filename in file_list:
            if filename.endswith(extension):
                pdf_list.append(os.path.join(dir_path, filename))
    return pdf_list

def extract_invoces(pdf_list, logger):
    """
    Get the invoices from the pdf files in th pdf_list
    Wrapper around the pdf2rawtxt call

    :param list pdf_list: List of pdf files path to process.
    :param logger: :class:`StatLogger`, collect statistical data about parsing

    :return: list of invoices
    :rtype: list of :class:`Invoice`
    """
    invoice_list = []
    for pdfile in pdf_list:
        invoice_list.append(pdf2rawtxt(pdfile, logger))
    return invoice_list

def _post_clean_up(tmp_dir='tmp'):
    """
    Cleanup after execution, remove the extracted zip file and tmp directory

    :param str tmp_dir: Temporary directory to clean_up
    """
    shutil.rmtree(tmp_dir)


def invoices2xlsx(invoices, directory='', name='Invoices01.xlsx'):
    """
    Write invoice information to xlsx template file. Go through every invoce and
    write them out. Simple. Utilizes the openpyxl module

    :param invoices list of Invocie: Representation of invoices from the pdf files
    """
    workbook = Workbook()
    worksheet_invo = workbook.active
    worksheet_entr = workbook.create_sheet()
    row_invo = col_invo = row_entr = col_entr = 0

    labels = ["Invoice Number", "Date of Invoice", "Payment Date", "Amount"]
    positions = config['invo_header_ident']['value']
    row_invo, col_invo = list2row(worksheet_invo, row_invo,
                                  col_invo, labels, positions)

    labels = ["Invoice Number"] + list(EntryTuple._fields)
    row_entr, col_entr = list2row(worksheet_entr, row_entr,
                                  col_entr, labels)
    #[TODO] there is no specification how to write out invocie entries yet
    for invo in invoices:
        row_invo, col_invo = invo.xlsx_write(worksheet_invo, row_invo, col_invo)
        for entr in invo.entries:
            row_entr, col_entr = entr.xlsx_write(worksheet_entr, row_entr, col_entr)
    #increase the invoice tab B column width to show the whole id
    worksheet_invo.column_dimensions['B'].width = 15

    workbook.save(os.path.join(directory, name))  

def run_excel(xlsx_path):
    """
    Start up Excel, with the file from the argument. The location of the excel
    executable should be set in the configuration

    :param str xlsx_path: Path to the xlsx file to open
    """
    run([config['excel_path']['value'], xlsx_path])


def do_it(src_name, dst_dir='', xlsx_name='Invoices01.xlsx',
          tmp_dir='tmp', file_extension='.pdf'):
    """
    Main script to manage the zip to xls process. It is responsible to create/cleanup
    temporary directories and files. After zip extraction, seraches every file which
    ends with `file_extension` Then it builds up a list of invoices and writes them
    to xlsx format and opens it up in the predefined xlsx_viewer. If the dst_dir happens
    to be the same as the tmp_dir, the generated xlsx file is removed after the run.

    :param str src_name: path to the zip file to extract
    :param str dst_dir: path to the directory to put the generated xlsx file by default
        the cwd
    :param str tmp_dir: temporary directory to work in. **This directory is erased
        at the beginning of the script** By default it is `tmp`
    :param str file_extension: the file extension to use during file selection. By
        default it is `.pdf`
    :param str xlsx_name: Name of the oputput file
    """
    _init_clean_up(tmp_dir)

    extract_zip(src_name, tmp_dir)

    pdf_list = get_pdf_files(os.path.join(os.getcwd(), tmp_dir), file_extension)

    logger = StatLogger()

    invoice_list = extract_invoces(pdf_list, logger)

    invoices2xlsx(invoice_list, dst_dir, name=xlsx_name)

    run_excel(os.path.join(dst_dir, config['xlsx_name']['value']))

    print(logger)

    print("script has been finished")
    return logger


def do_it2(src_name, dst_dir='', xlsx_name='Invoices01.xlsx',
           tmp_dir='tmp'):
    """
    Main script to manage the zip to xls process. It is responsible to create/cleanup
    temporary directories and files. After zip extraction, seraches every file which
    ends with `file_extension` Then it builds up a list of invoices and writes them
    to xlsx format and opens it up in the predefined xlsx_viewer. If the dst_dir happens
    to be the same as the tmp_dir, the generated xlsx file is removed after the run.

    :param str src_name: path to the zip file to extract
    :param str dst_dir: path to the directory to put the generated xlsx file by default
        the cwd
    :param str tmp_dir: temporary directory to work in. **This directory is erased
        at the beginning of the script** By default it is `tmp`
    :param str xlsx_name: Name of the oputput file
    """
    _init_clean_up(tmp_dir)

    order_list = read_xlsx(GetOrderDetail, src_name)

    write_xlsx(order_list, filename=os.path.join(dst_dir, config['xlsx_name']['value']))

    run_excel(os.path.join(dst_dir, config['xlsx_name']['value']))

    print("Order detail extraction script has been finished")
    return True


def main():
    """
    A simple wrapper around do it function, to demonstrate its usage
    """
    do_it(src_name='src.zip', dst_dir='', xlsx_name=config['xlsx_name']['value'],
          tmp_dir=config['tmp_dir'][0], file_extension=config['file_extension']['value'])


if __name__ == '__main__':
    main()
