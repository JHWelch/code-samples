'''
Script to download all attachments on a Smartsheet labeled by row using the Smartsheet API

https://smartsheet-platform.github.io/api-docs/

https://github.com/smartsheet-platform/smartsheet-python-sdk
'''
import smartsheet
import logging
import re
import time
import smtplib
import os
import urllib

# Set logging to default config so that smartsheet errors are logged corectly
logging.basicConfig()


#Set constants to the correct IDs of the sheet being downloaded, and add your API key
OUTPUT_FOLDER = "C:\\"

LOG_FILENAME = "Smartsheet_download.log"

DEFAULT_CSV_NAME = "Node Handoff Document.csv"

SMARTSHEET_API_KEY = 'SECRET_API_KEY'

SMARTSHEET_ID  = 'ID GOES HERE'

FILE_NAME_COL_ID = 'COLUMN ID GOES HERE'

def WriteLogLine(Line):
    with open(OUTPUT_FOLDER + LOG_FILENAME, 'a') as d:
        d.write(time.strftime('%Y-%m-%d %H:%M:%S') + "|" + Line + "\n")

def WriteLogLines(Lines = []):
    with open(OUTPUT_FOLDER + LOG_FILENAME, 'a') as d:
        for line in Lines:
            d.write(time.strftime('%Y-%m-%d %H:%M:%S')+ "|" + line + "\n")

def getCellByColumnId(row, columnId):
    return next(cell for cell in row.cells if cell.column_id == columnId)

ss_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)

WriteLogLine("Starting Script Execution")

sheet_input = ss_client.Sheets.get_sheet(SMARTSHEET_ID, page_size = 10000,include = "attachments")

'''
Iterate through rows in the input sheet, filling input row and origin row arrays
that will be used to update individual sheets and delete from the Input sheet
'''
WriteLogLine("Reading from Origin Sheet")
for row in sheet_input.rows:
    attach_count = len(row.attachments)
    if attach_count > 0:
        cur_row = 1
        for attachment in row.attachments:
            suffix = "";
			# If there is more than 1 attachment, add suffix indicating which attachment it is.
            if attach_count > 1:
                suffix = "-" + str(cur_row) + "-of-" + str(attach_count)
            attach_to_download = ss_client.Attachments.get_attachment(SMARTSHEET_ID, attachment.id)

            address = getCellByColumnId(row, ID_PROPERTY_ADDRESS)

            urllib.urlretrieve(attach_to_download.url, OUTPUT_FOLDER + address.value + suffix + ".jpeg")
            print "Saved: " + address.value + suffix + ".jpeg"
            WriteLogLine("Saved: " + address.value + suffix + ".jpeg")
            cur_row = cur_row + 1
            