#!/usr/bin/env python3
# -*- coding: utf-8 -*-

##############################################################################
# This script:
# get table data from first (zero-indexed) sheet of *.xlsx file,
# put it into dictionary, 
# convert to bytes by pickle
# and save *.pkl bytes file to EXPORT_DIR folder.
##############################################################################

from pathlib import Path
import logging
from datetime import datetime
import pickle

# Set path to log file
SCRIPT_DIR = Path(".config/alteroffice/5/user/Scripts/python")
LOG_FILE = "export_dictionary.log"
LOG_PATH = Path.home() / SCRIPT_DIR / LOG_FILE

# Change a log level like below for logging in LOG_FILE
# logging.basicConfig(filename=LOG_PATH, level=logging.INFO)
logging.basicConfig(filename=LOG_PATH, level=logging.ERROR)

logging.info("Start `export_dictionary.py` script. See listing below:")

# Set export folder
EXPORT_DIR = "/full/path/to/export/folder/"

def get_context_data() -> dict:
    """
    Get the data from table and return dictionary,
    where first row (labels) is a keys.

    From *.xlsx table:
    |k1 | k2|
    ---------
    |v1 | v2|
    |v1 | v2|
    |v1 | v2|

    To dictionary:

    {
        k1 : [v1,v1,v1],
        k2 : [v2,v2,v2]
    }
    """
    logging.info("Start `get_context_data` function.")

    logging.info("Try to get data from *.xlsx table...")
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)
    cols: int = len(sheet.ColumnDescriptions)
    rows: int = len(sheet.RowDescriptions)

    logging.info("Try to put data into dictionary...")
    data: dict = dict()

    for i in range(cols):
        for j in range(rows):
            if j==0:
                key = sheet.getCellByPosition(i,j).String
                data[key]: list = []
            else:
                data[key].append(sheet.getCellByPosition(i,j).String)

    logging.info("The data was writed into dictionary!")
    return data


def get_data_as_bytes(data: dict):
    """
    Get data from *.xlsx file,
    and convert it to bytes by pickle.
    """
    logging.info("Start `get_data_as_bytes` function.")

    logging.info("Try to convert data to bytes...")
    buffer = pickle.dumps(data)

    logging.info("The data was converted to bytes!")
    return buffer

def export_dictionary():
    """
    Write the `yyyy-mm-dd_hh-mm-ss_ms.pkl` file
    into `EXPORT_DIR` folder.
    """
    logging.info("Start `export_data` function.")

    logging.info("Set a current datetime into local variable...")
    dts = str(datetime.now())

    data = get_context_data()
    buffer = get_data_as_bytes(data)

    logging.info("Construct file name from datetime of data")
    s = [EXPORT_DIR, dts[0:10], "_", dts[11:19], "_", dts[20:], ".pkl"]
    file_name = "".join(s).replace(":","-")

    logging.info("Try to write data into *.pkl file...")
    with open(file_name,"wb") as file:
        file.write(buffer) 
        logging.info(f"Data was succesfully writed into `{file_name}` file!")

    return
