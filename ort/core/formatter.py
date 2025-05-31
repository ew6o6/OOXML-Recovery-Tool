"""
core/formatter.py - Save extracted data into document or table format
role: Module for handling extracted data from OOXML files

author: Jiyoon Kim
last-updated: 2025-05-12

description:
    extract_xml_text_for_docx() - Extracts text from OOXML document XML
    save_text_to_docx() - Saves extracted text to a .docx file
    extract_data_from_shared_strings() - Extracts data from sharedStrings.xml
    extract_data_from_sheet() - Extracts data from worksheet XML
    save_unmapped_to_csv() - Saves unmatched data to CSV
    display_and_save_table_to_csv() - Saves and displays matched table data to CSV
    process_extracted_docx_data() - Handles and saves extracted DOCX data
    process_extracted_xlsx_data() - Handles and saves extracted XLSX data
    process_extracted_data() - Handles and saves extracted OOXML data depending on file extension
"""
import os
import re
import csv
from bs4 import BeautifulSoup
from tabulate import tabulate
from docx import Document

def parse_xlsx_rels_file(rels_xml):
    """Parses relationships in XLSX file and returns relation ID, target, and type"""
    soup = BeautifulSoup(rels_xml, "xml")
    rels_info = []
    for rel in soup.find_all("Relationship"):
        r_id = rel.get("Id")
        target = rel.get("Target")
        r_type = rel.get("Type")
        rels_info.append(f"{r_id}: {target} ({r_type})")
    return rels_info

def process_extracted_data(local_file_xml, file_path, ext):
    """Dispatch function that processes extracted data based on OOXML extension"""
    if ext == '.docx':
        process_extracted_docx_data(local_file_xml, file_path)
    elif ext == '.xlsx':
        process_extracted_xlsx_data(local_file_xml, file_path)
    else:
        print("Unsupported file format.")
