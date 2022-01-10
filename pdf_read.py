import PyPDF2
import logging
import re


def write_worksheet_row(worksheet, row, content):
    col = 0
    for item in content:
        worksheet.write(row, col, item)
        col += 1
    row += 1
    return row


def search_by_page(file_path, worksheet, worksheet_row, word):

    # read the pdf file
    pdf_file = open(file_path, 'rb')
    logging.debug("Read file" + file_path)

    if pdf_file == None:
        # File not found
        logging.error("File does not exist")
        return

    pdf_reader = PyPDF2.PdfFileReader(pdf_file)
    logging.info(file_path + " has " + str(pdf_reader.numPages) + " pages")

    # Iterate through pages
    for page_no in range(0, pdf_reader.numPages):
        page_obj = pdf_reader.getPage(page_no)
        page_text = page_obj.extractText()
        words = re.findall(word, page_text, re.IGNORECASE)
        worksheet_content = []
        worksheet_content.append(file_path)
        worksheet_content.append(str(page_no))
        worksheet_content.append(word)
        worksheet_content.append(str(len(words)))
        worksheet_row = write_worksheet_row(
            worksheet=worksheet, row=worksheet_row, content=worksheet_content)
        logging.info("page no. {}, freq {}", page_no, len(words))
    pdf_file.close()
    return worksheet_row
