# OOXML Recovery Toolkit (ORT)

A lightweight CLI tool to recover text, images, and metadata from corrupted OOXML files (`.docx`, `.xlsx`, `.pptx`).

**GitHub Repository:**  
https://github.com/ew6o6/OOXML-Recovery-Tool

## Features

- Supports damaged Word, Excel, and PowerPoint files
- Extracts text content, images, and relationships
- Outputs recovered artifacts into a simple folder structure

## Requirements

- Python 3.8+
- Dependencies listed in `setup.py` (e.g., `beautifulsoup4`, `lxml`, `python-docx`)

## Installation

1. **Clone the repo**
   '''bash
   git clone https://github.com/ew6o6/OOXML-Recovery-Tool
   cd <REPO-DIR>
   '''
2. **(Optional) Create and activate a virtual environment**
   '''bash
   python3 -m venv .venv
   source .venv/bin/activate
   ''' 3.**Install in editable mode**
   pip install --upgrade pip
   pip install .

## Usage

1. **As an installed script**
   '''

   # Recover a single file

   ort path/to/file.docx

   # Recover all OOXML files in a directory

   ort path/to/some_folder/

   # Recover a PowerPoint file

   ort ./slides/broken.pptx
   '''

2. **As a Python module**
   '''
   # From the project root (where setup.py lives):
   python -m ort path/to/file.xlsx
   python -m ort path/to/directory/
   '''

## Output Structure (Example)

    '''
    output_damaged/
    ├── img/
    ├── metadata.json
    ├── recovered.docx
    └── relationships.txt
    '''
