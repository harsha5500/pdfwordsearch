import pdf_read
import logging
import xlsxwriter

logging.basicConfig(level=logging.INFO)

# create a xlsx file
HEADING_ROW = ["File", "Page No", "Word", "Word Frequency"]
SUMMARY_FILE = "summary.xlsx"

workbook = xlsxwriter.Workbook(SUMMARY_FILE)
worksheet = workbook.add_worksheet()
worksheet_row = 0

worksheet_row = pdf_read.write_worksheet_row(
    worksheet=worksheet, row=worksheet_row, content=HEADING_ROW)

# Test open lshort file
FILE = "./sample.pdf"

worksheet_row = pdf_read.search_by_page(FILE, worksheet, worksheet_row, "text")

workbook.close()
