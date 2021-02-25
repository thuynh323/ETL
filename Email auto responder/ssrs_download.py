import win32cred
import os
import glob2
import configparser
import re
from openpyxl import Workbook
from requests import Session
from requests_ntlm2 import HttpNtlmAuth
import win32com.client


outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
email_address = outlook.DefaultStore
username = f"{os.environ.get('USERDOMAIN')}\\{os.environ.get('USERNAME')}"
cred_target = f'Microsoft_OC1:uri={email_address}:specific:CER:1'
credential = win32cred.CredRead(cred_target, win32cred.CRED_TYPE_GENERIC, 0)
password = credential.get('CredentialBlob').decode('utf-16')

config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())
for config_file in glob2.glob('*.txt'):
    config.read(config_file)
    path = config['BASE']['path']
    render = config['RENDER']['render']
    session = Session()
    session.auth = HttpNtlmAuth(username=username,password=password)
    if config['BASE'].getboolean('multi_sheets') == False:
        for (url_key, url_value), (report_key, report_value) in zip(config.items('URL'), config.items('REPORT_NAME')):
            url = url_value + render
            response = session.get(url)
            report_name = report_value
            file_path = os.path.join(path, f'{report_name}.xlsx')
            with open(os.path.join(path, file_path), 'wb') as report_excel:
                for chunk in response.iter_content(chunk_size=100000):
                    report_excel.write(chunk)
    else:
        for (url_key, url_value), (sheet_key, sheet_value) in zip(config.items('URL'), config.items('SHEET_NAME')):
            url = url_value + render
            response = session.get(url)
            report_name = sheet_key
            file_path = os.path.join(path, f'{report_name}.xlsx')
            with open(os.path.join(path, file_path), 'wb') as report_excel:
                for chunk in response.iter_content(chunk_size=100000):
                    report_excel.write(chunk)
        combined_report_name = config['RENDER']['report_name']
        workbook = Workbook()
        workbook_path = os.path.join(path, f'{combined_report_name}.xlsx')
        workbook.save(workbook_path)
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = True
        for wb in glob2.glob(os.path.join(path,'*[0-9].xlsx')):
            parent_wb = excel.Workbooks.Open(workbook_path)
            child_ws = excel.Workbooks.Open(wb).Worksheets(1)
            child_ws.Name = config['SHEET_NAME'][re.search(r'(?:sheet_[0-9])', wb).group(0)]
            child_ws.Move(Before=parent_wb.Worksheets(1))
            parent_wb.Close(SaveChanges=True)
            excel.Quit()
            os.remove(wb)       

    session.close()