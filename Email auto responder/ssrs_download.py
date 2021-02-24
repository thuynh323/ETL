import win32cred
import os
import configparser
import json
from requests import Session
from requests_ntlm2 import HttpNtlmAuth
import win32com.client

outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
email_address = outlook.DefaultStore
username = f"{os.environ.get('USERDOMAIN')}\\{os.environ.get('USERNAME')}"
cred_target = f'Microsoft_OC1:uri={email_address}:specific:CER:1'
credential = win32cred.CredRead(cred_target, win32cred.CRED_TYPE_GENERIC, 0)
password = credential.get('CredentialBlob').decode('utf-16')

param = json.load(open(r'path\config.json'))
base_url = param.get('url')
path = param.get('path')
report_path = param.get('report')

url = base_url + report_path + '&rs:Format=EXCELOPENXML'
session = Session()
session.auth = HttpNtlmAuth(username=username,password=password)
response = session.get(url)

with open('report.xlsx', 'wb') as report_excel:
    for chunk in response.iter_content(chunk_size=100000):
        report_excel.write(chunk)
session.close()