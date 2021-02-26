import requests
import win32com.client as win32
import configparser
import json
import re
from datetime import datetime
from os import environ
from pathlib import Path

def track_pic(url: str,
              headers: str,
              ca_path: str,
              tracking_number: str) -> dict:
    """
    Returns a dictionary of tracking summary or tracking error
    """
    r = requests.get(f'{url}{tracking_number}',
                     headers=headers,
                     verify=ca_path)
    response = r.json().get('trackResponse').get('shipment')[0].get('package')

    if response == None:
        tracking_summary = {}
    else:
        tracking_summary = response[0].get('activity')[0]
    return tracking_summary

def tracking_details(tracking_number: str,
                     tracking_summary: dict) -> str:
    """
    Returns mail content
    """
    
    if tracking_summary == {}:
        mail_body = f"""
            <html>
            <title></title>
            <body>
                <font face='Calibri' size='-0.5'>
                <p>Hello,</p>
                <p>Unfortunately, we are unable to locate this package in our system.</p>
                </font>
            </body>
            </html>
            """
    else:
        location = tracking_summary.get('location').get('address')
        status = tracking_summary.get('status')

        last_city = location.get('city')
        last_state = location.get('stateProvince')
        last_zip = location.get('postalCode')
        last_status = status.get('description')
        last_date = datetime.strptime(tracking_summary.get('date'), '%Y%m%d').date()
        last_time = datetime.strptime(tracking_summary.get('time'), '%H%M%S').time()
    
        if last_city != None:
            mail_body = f"""
                <html>
                <title></title>
                <body>
                    <font face='Calibri' size='-0.5'>
                    <p>Hello,</p>
                    <p>Your request has been received and is being reviewed by our support department. 
                    While we investigate this package, we have set up an email alert with USPS for you
                    to receive updates until the package is delivered.</p>
                    <p>Tracking number:<br>
                    &emsp;{tracking_number}<br>
                    Current package status:<br>
                    &emsp;{last_status}&emsp;{last_date} {last_time}<br>
                    USPS Entry Point:<br>
                    &emsp;{last_city}, {last_state} {last_zip}</p>
                    </font>
                </body>
                </html>
                """
        else:
            mail_body = f"""
                <html>
                <body>
                    <font face='Calibri' size='-0.5'>
                    <p>Hello,</p>
                    <p>Your request has been received and is being reviewed by our support department. 
                    While we investigate this package, we have set up an email alert with USPS for you
                    to receive updates until the package is delivered.</p>
                    <p>Tracking number:<br>
                    &emsp;{tracking_number}<br>
                    Current package status:<br>
                    &emsp;{last_status}</p>
                    </font>
                </body>
                </html>
                """
    return mail_body

def reply_mail(mail_items: object,
               url: str,
               headers: str,
               ca_path: str):
    """
    Read mail items and extract tracking numbers from the content
    Draft response emails including the latest tracking event retrieved from API
    """
    # Iterarte over mail items
    for mail in mail_items:
        # Select emails only
        if '_MailItem' in str(type(mail)):
            mail_content = mail.Body
            # Extract tracking number
            find_tracking_number = re.search(r'92\d{24}', mail_content)
            if find_tracking_number != None:
                tracking_number = find_tracking_number.group(0)
                tracking_result = track_pic(url=url,
                                            headers=headers,
                                            ca_path=ca_path,
                                            tracking_number=tracking_number)

                mail_body = tracking_details(tracking_number=tracking_number,
                                             tracking_summary=tracking_result)
                # Draft response in HTML format
                reply_all = mail.ReplyAll()
                reply_all.HTMLBody = mail_body + reply_all.HTMLBody
                reply_all.Save()
            else:
                continue
        else:
            continue

def main():
    
    # Read in the config file
    config = configparser.ConfigParser(interpolation=configparser.BasicInterpolation())
    config.read('config.ini')
    
    email_account = config['DEFAULT']['email_to_read']
    api_key = config['DEFAULT']['api_key']
    headers = {'AccessLicenseNumber': api_key}
    ca_path = config['DEFAULT']['ca_path']
    url = config['DEFAULT']['base_url']
    main_folder = config['DEFAULT']['main_folder']
    sub_folder = config['DEFAULT']['sub_folder']
    
    # Access Outlook
    try:
        outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    except AttributeError:
        username = environ.get('USERNAME')
        file_loc = f"C:\\Users\{username}\\AppLocal\\Temp\\gen_py"
        for f in file_loc:
            Path.unlink(f)
        Path.mkdir(f_loc)
        outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    # Select email and folder to read    
    account = outlook.Folders[email_account]
    folder = account.Folders[main_folder]
    read_folder = folder.Folders[sub_folder]
    
    # Draft responses
    reply_mail(mail_items=read_folder.Items,
               url=url,
               headers=headers,
               ca_path=ca_path)

if __name__ == '__main__':
    main()