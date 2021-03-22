import os
import openpyxl
import gspread
import google.auth
from google.oauth2 import service_account

def main():
    print('Fill tables greets u!')
    fill_table()
    #fill_table(reports)
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
def fill_table():
    gc = gspread.service_account() # Вход в аккаунт гугл бота
    table = gc.open('TestTable') # Открываем гугл-таблицу
    print("Successfull authorized")
    reports = load_reports()
    for report in reports:
        fill_month(table,report)
        fill_Pertrader_month(table,report)
        fill_balances(table,report)
    print('Tables loaded')
    #print("Table filled")
def fill_month(table,report):
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
    sheet_month = table.worksheet(numbers_to_string_month[report.month] + ' ' + report.year)
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
def fill_Pertrader_month(table,report):
    print('PerTrader ' + '' + ' filled')
def fill_balances(table,report):
    print('Balances filled')
#Вспомогательная функция разбиения имени отчёта? для получения инфы по нему
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
   

if __name__ == '__main__':
    main()