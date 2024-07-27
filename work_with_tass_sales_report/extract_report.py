import openpyxl
from bs4 import BeautifulSoup


# check file extension and extract data from suitable file
def extract_mail_report(file_extension, path_to_report_file):
    mail_report = None
    if file_extension == '.html':
        # for the html report
        mail_report = report_from_html_report_file(path_to_report_file)
    elif file_extension == '.xlsx':
        # for the xlsx report
        mail_report = report_from_tass_xlsx_file(path_to_report_file)
    else:
        print("wrong report file type")
    # index row number, value - list from columns date
    return mail_report


# from html report file extract dict with information about sales,
# index row number, value - list from columns date
def report_from_html_report_file(path_to_report_file: str) -> dict:
    with open(path_to_report_file, 'r') as report_file:
        table = BeautifulSoup(report_file, 'lxml')
    table = table.find('tbody')
    table = table.find_all('tr')[4:]  # start with data row in table
    row_in_table = [x for x in table]
    report = {}
    for i in range(len(table)):
        report[i] = [x.text.strip() for x in row_in_table[i].find_all('td')]
    return report


# create dict from xlsx file, index row number, value - list from columns date
def report_from_tass_xlsx_file(xlsx_file: str) -> dict:
    wb = openpyxl.load_workbook(xlsx_file)
    sheet = wb.active
    report = {}
    for number, value in enumerate(sheet.rows, start=1):
        report[number] = [cell.value if cell.value is not None else '' for cell in sheet[number]]
    return report
