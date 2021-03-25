import os
import openpyxl
import pandas as pd
import requests
import google.auth
from google.oauth2 import service_account
from gspread_pandas import Spread, Client 

#Основная функция 
def main():
    print('Fill tables greets u!')
    fill_table()
# Функция загрузки отчётов из таблиц
def load_reports():
    reports = []
    reports_counter = 0
    reports_folder = os.listdir(os.getcwd() + '/reports')
    for report_file in reports_folder:
        if report_file[0] != '.':
            workbook = openpyxl.load_workbook(filename = "./reports/" + report_file)
            reports.append(Report(str(report_file),workbook))
            reports_counter += 1
    print("Reports loaded")
    return reports
# Функция заполнения таблиц
def fill_table():
    try:
        table = Spread('PnL New')
        print("Successfull authorized")
    except requests.exceptions.RequestException:
        print("Can't connect to Google API")
        raise SystemExit()
    reports = load_reports()
    for report in reports:
        fill_month(table,report)
        fill_Pertrader_month(table,report)
        fill_balances(table,report)
    print("Table filled")
# Заполнение листа месяца
def fill_month(table,report):
    sheet_month = table.worksheet(get_month_from_numbers(report.month) + ' ' + report.year)
    filling_row = sheet_month.find(report.get_date()).row
    if (report.session == 'B'):
        filling_col = sheet_month.find(report.burse).col 
    else:
        filling_col = sheet_month.find(report.burse + ' Night').col 
    col_counter = 1
    cell_value = report.workbook.get_sheet_by_name("Total").cell(2,col_counter).value
    while cell_value != None:
        sheet_month.update_cell(filling_row, filling_col + col_counter - 1, cell_value)
        col_counter += 1
        cell_value = report.workbook.get_sheet_by_name("Total").cell(2,col_counter).value
    print('Month filled')
#Заполнение листа трейдеров
def fill_Pertrader_month(table,report):

    #Определение диапазона копируемых ячеек
    print('Works on PerTrader sheet')
    report_sheet = report.workbook.get_sheet_by_name('Per_trader')
    first_report_row = 2
    first_report_col = 1
    last_report_row = report_sheet.max_row
    last_report_col = first_report_col + 4
    copied_range_str = range_to_a1(first_report_row, first_report_col, last_report_row, last_report_col)
    copied_range = report_sheet[number_to_fucking_a1(first_report_col) + str(first_report_row):number_to_fucking_a1(last_report_col) + str(last_report_row)]
    flat_copied_range = []
    for tup in copied_range:
        for t in tup:
            flat_copied_range.append(t.value)

    #Определение диапазона заполняемых ячеек
    sheet_PerTrader = table.find_sheet('PerTrader_' + str(get_month_from_numbers(report.month)[0:3]))
    first_filling_col = sheet_PerTrader.find('session').col
    session_column = sheet_PerTrader.col_values(first_filling_col)
    first_filling_row = len(session_column) + 1
    last_filling_col = first_filling_col + last_report_col - first_report_col
    last_filling_row = first_filling_row + last_report_row - first_report_row
    filling_range = range_to_a1(first_filling_row, first_filling_col, last_filling_row, last_filling_col)
    print('Copying {} from report_sheet to {} in PerTrader_sheet.'.format(copied_range_str,filling_range))
    #print('Number of values = ' + str(len(flat_copied_range)))
    #print('Number of cells = ' + str((last_filling_col - first_filling_col + 1) * (last_filling_row - first_filling_row + 1)))

    # Заполнение таблицы
    table.update_cells((first_filling_row,first_filling_col),(last_filling_row,last_filling_col),list(flat_copied_range),sheet_PerTrader)
#Заполнение листа балансов
def fill_balances(table,report):

    #Определение диапазона копируемых ячеек
    print('Works on Balances sheet')
    report_sheet = report.workbook.get_sheet_by_name('Balances')
    first_report_row = 2
    first_report_col = 1
    last_report_row = report_sheet.max_row
    last_report_col = first_report_col + 5
    copied_range_str = range_to_a1(first_report_row, first_report_col, last_report_row, last_report_col)
    copied_range = report_sheet[number_to_fucking_a1(first_report_col) + str(first_report_row):number_to_fucking_a1(last_report_col) + str(last_report_row)]
    flat_copied_range = []
    for tup in copied_range:
        for t in tup:
            flat_copied_range.append(t.value)

    #Определение диапазона заполняемых ячеек
    sheet_Balances = table.find_sheet('Balances')
    first_filling_col = 1
    session_column = sheet_Balances.col_values(first_filling_col)
    first_filling_row = len(session_column) + 1
    last_filling_col = first_filling_col + last_report_col - first_report_col
    last_filling_row = first_filling_row + last_report_row - first_report_row
    filling_range = range_to_a1(first_filling_row, first_filling_col, last_filling_row, last_filling_col)
    print('Copying {} from report_sheet to {} in PerTrader_sheet.'.format(copied_range_str,filling_range))

    # Заполнение таблицы
    table.update_cells((first_filling_row,first_filling_col),(last_filling_row,last_filling_col),list(flat_copied_range),sheet_Balances)

    print('Balances filled')
#Вспомогательная функция разбиения имени отчёта, для получения инфы по нему
def parse_report_name(report_file):
    current_year = report_file[:4]
    current_month = report_file[4:6]
    current_day = report_file[6:8]
    burse = ''
    counter = 9
    while report_file[counter] != '_':
        burse += report_file[counter]
        counter += 1
    session = report_file[counter + 1:counter + 2]
    print('Year ' + current_year + ' Month ' + current_month + ' day ' + current_day + ' burse ' + burse + ' session ' + session)
    print('Report name successfully parsed')
    return current_year, current_month, current_day, burse, session
#Класс отчёта
class Report():
    year = 0
    month = 0
    day = 0
    burse = ''
    session = ''
    workbook = openpyxl.Workbook()
    def __init__(self,report_name,workbook):
        self.year, self.month, self.day, self.burse, self.session = parse_report_name(report_name)
        self.workbook = workbook
    def get_date(self):
        return str(self.year + '.' + self.month + '.' + self.day)
#Получение строки месяца по числу
def get_month_from_numbers(number):
    numbers_to_string_month = {
        '01' : 'January',
        '02' : 'February',
        '03' : 'March',
        '04' : 'April',
        '05' : 'May',
        '06' : 'June',
        '07' : 'July',
        '08' : 'August',
        '09' : 'September',
        '10' : 'October',
        '11' : 'November',
        '12' : 'December'
    }
    return numbers_to_string_month[number]
#Поиск первой пустой строки в таблице
def get_empty_raw(session_col):
    str_list = filter(None,session_col)
    return len(str_list) + 1
# Приведение нотации таблицы к формату A1
def number_to_fucking_a1(number):
    number_to_a1 = {
        0 : '',
        1 : 'A',
        2 : 'B',
        3 : 'C',
        4 : 'D',
        5 : 'E',
        6 : 'F',
        7 : 'G',
        8 : 'H',
        9 : 'I',
        10 : 'J',
        11 : 'K',
        12 : 'L',
        13 : 'M',
        14 : 'N',
        15 : 'O',
        16 : 'P',
        17 : 'Q',
        18 : 'R',
        19 : 'S',
        20 : 'T',
        21 : 'U',
        22 : 'V',
        23 : 'W',
        24 : 'X',
        25 : 'Y',
        26 : 'Z'
    }
    a1_notation = ''
    while number // 26 > 0:
        a1_notation += number_to_a1[number // 26]
        number = number % 26
    a1_notation += number_to_a1[number]
    return a1_notation
# Преобразование диапазона ячеек к А1 нотации
def range_to_a1(first_row, first_col, last_row, last_col):
    #start = number_to_fucking_a1(first_col) + str(first_row)
    #end = number_to_fucking_a1(last_col) + str(last_row)
    range = number_to_fucking_a1(first_col) + str(first_row) + ':' + number_to_fucking_a1(last_col) + str(last_row)
    return range
if __name__ == '__main__':
    main()