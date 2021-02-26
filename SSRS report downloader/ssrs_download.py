import win32cred
import win32com.client
import os
import configparser
import glob2
import re
from openpyxl import Workbook
from requests import Session
from requests_ntlm2 import HttpNtlmAuth

def download_data(username: str,
                  password: str,
                  combine: bool,
                  path: str,
                  url_section: list,
                  name_section: list):
    
    """
    Downloads data from an SSRS URL and writes them into an excel file
    then store the file in the given path

    username -- your username in SSRS
    password -- your login password (same as your PC password)
    combine -- if you want to combine these reports
    path -- directory to save the reports
    url_section -- list of (key, value) in the config file, under URL
    name_section -- list of (key, value) in the config file, under REPORT_NAME or REPORT_NAME
    """
    
    # Open a session
    session = Session()
    session.auth = HttpNtlmAuth(username=username,password=password)
    
    # Iterate over each pair of url and name
    for (url_key, url_value), (name_key, name_value) in zip(url_section, name_section):

        response = session.get(url_value)
        # Create a report excel file
        # Name it based on combine option
        if combine:
            report_name = name_key
        else:
            report_name = name_value

        # Write in binary to keep the format 
        file_path = os.path.join(path, f'{report_name}.xlsx')
        with open(file_path, 'wb') as report_excel:
            for chunk in response.iter_content(chunk_size=100000):
                report_excel.write(chunk)

    # Close the session
    session.close()

def combine_excel(report_name: str,
                  path: str,
                  sheet_names: dict):

    """
    Combines multiple one-sheet workbooks to one file that
    has multiple sheets and store it in the given path

    report_name -- name of the big report
    path -- directory to save the report
    sheet_names -- pairs of a child report name before combining and
                   its sheet name after combining 
    """

    # Create a blank workbook. This is the parent workbook
    # Name it as the report name and save it to the given path
    workbook = Workbook()
    workbook_path = os.path.join(path, f'{report_name}.xlsx')
    workbook.save(workbook_path)

    # Open Excel silently
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    
    # In the given path,
    # iterate over all excel files have name pattern sheet_1, sheet_2, and so on.
    # These are child workbooks
    for wb in glob2.glob(os.path.join(path,'*[0-9].xlsx')):
        
        # Open the parent workbook again
        parent_wb = excel.Workbooks.Open(workbook_path)
        # Open current child workbook
        # Select the sheet 
        child_ws = excel.Workbooks.Open(wb).Worksheets(1)

        # Rename this sheet properly by matching child report's name 
        # and its new worksheet name defined in the config file under SHEET_NAME
        child_ws.Name = sheet_names[re.search(r'(?:sheet_[0-9])', wb).group(0)]

        # Move the renamed sheet to the parent workbook
        child_ws.Move(Before=parent_wb.Worksheets(1))
        # Save and close the parent workbook
        parent_wb.Close(SaveChanges=True)
        # Close Excel
        excel.Quit()
        
        # Remove the child report from the directory
        os.remove(wb)

def main():
    
    # Get your domain and username 
    username = f"{os.environ.get('USERDOMAIN')}\\{os.environ.get('USERNAME')}"

    # Get your default email which is linked to your cedential
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    email_address = outlook.DefaultStore
    # Get the credential 
    cred_target = f'Microsoft_OC1:uri={email_address}:specific:CER:1'

    # Your credential might be different
    # Another option is to create a separate config file for your credential

    # Now get your login password
    credential = win32cred.CredRead(cred_target, win32cred.CRED_TYPE_GENERIC, 0)
    password = credential.get('CredentialBlob').decode('utf-16')

    # Set up configuration parser
    config = configparser.ConfigParser(interpolation=configparser.ExtendedInterpolation())
    
    # Iterate through all .ini files in the current directory
    for config_file in glob2.glob('*.ini'):

        # Read and extract values from the config file
        config.read(config_file)
        path = config['BASE']['path']
        combine_option = config['REPORT_OPTION'].getboolean('combine')
        url_section = config.items('URL')
        report_name_section = config.items('REPORT_NAME')
        sheet_name_section = config.items('SHEET_NAME')
        
        # Execute based on combine option
        if combine_option:
            # Combine is True -> write reports and name them by sheet keys defined in the config file
            # under SHEET_NAME
            download_data(username=username,
                          password=password,
                          combine=combine_option,
                          path=path,
                          url_section=url_section,
                          name_section=sheet_name_section)
            # Get name of the filnal report
            combined_report_name = config['REPORT_OPTION']['report_name']
            # Get sheet names
            sheet_names = config['SHEET_NAME']
            # Combine the reports
            combine_excel(report_name=combined_report_name,
                          path=path,
                          sheet_names=sheet_names)
        else:
            # Combine is False -> write reports and name them by report names defined in the config file
            # under REPORT_NAME
            download_data(username=username,
                          password=password,
                          combine=combine_option,
                          path=path,
                          url_section=url_section,
                          name_section=report_name_section)

if __name__ == '__main__':
    main()