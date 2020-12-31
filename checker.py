import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import threading
import time

FILE_PATH = 'Data File.xlsx'
wb = load_workbook(FILE_PATH)


# takes an url and returns its status code and final redirected url
def get_response_code(url):
    url = 'https://' + url if not url.startswith('http') else url
    response = requests.get(url)
    final_url = response.url
    status_code = response.status_code
    if response.history:
        status_code = response.history[0].status_code
    return (url, final_url, status_code)


def customize_excel_sheet():
    global wb
    output = wb.create_sheet('Output') if 'Output' not in wb.sheetnames else wb['Output']
    
    font = Font(color="000000", bold=True)
    bg_color = PatternFill(fgColor='E8E8E8', fill_type='solid')

    # editing the output sheet
    output_column = zip(('A',  'B', 'C'), ('URL', 'Final URL', 'HTTP Response'))
    for col, value in output_column:
        cell = output[f'{col}1']
        cell.value = value
        cell.font = font
        cell.fill = bg_color
        output.freeze_panes = cell

        # fixing the column width
        output.column_dimensions[col].width = 20


# Generates the input links
def generate_input_urls():
    global wb
    inputs = wb['Input']
    for row in range(2, inputs.max_row + 1):
        # generates the links one by one
        if value := inputs[f"A{row}"].value:
            yield value


def insert_data_to_excel():
    global wb
    def save(url):
        data = get_response_code(url)
        wb['Output'].append(data)
        print(data)

    threads = []
    for url in generate_input_urls():
        thread = threading.Thread(target=save, args=[url])
        thread.start()
        threads.append(thread)
        
    
    for thread in threads:
        thread.join()

t1 = time.perf_counter()
thread = customize_excel_sheet()
insert_data_to_excel()
wb.save(FILE_PATH)
t2 = time.perf_counter()

print(t2 -t1)
