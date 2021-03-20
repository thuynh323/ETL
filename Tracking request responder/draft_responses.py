import requests
import win32com.client as win32
import configparser
import socket
import json
import re
from urllib import parse
from os import environ
from shutil import rmtree
from datetime import datetime

def track_pic(url: str,
              headers: str,
              ca_path: str,
              tracking_number: str) -> dict:
    """
    Returns a dictionary of tracking summary
    Returns empty dict if the tracking number is invalid
    """
    r = requests.get(f'{url}{tracking_number}', headers=headers, verify=ca_path)
    response = r.json().get('trackResponse').get('shipment')[0].get('package')
    
    tracking_events = {}
    if response == None:
        pass
    else:
        tracking_summary = response[0].get('activity')[0]
        location = tracking_summary.get('location').get('address')
        status = tracking_summary.get('status')

        last_status = status.get('description')
        last_date = datetime.strptime(tracking_summary.get('date'), '%Y%m%d').date()
        last_time = datetime.strptime(tracking_summary.get('time'), '%H%M%S').time()
        last_city = location.get('city')
        last_state = location.get('stateProvince')
        last_zip = location.get('postalCode')
        
        tracking_events['tracking_number'] = tracking_number
        tracking_events['last_event'] = last_status
        if last_city != None:
            tracking_events['last_location'] = last_city + ', ' + last_state + ' ' + last_zip
        if last_date != None:
            tracking_events['last_date_time'] = str(last_date) + ' ' + str(last_time)
    return tracking_events

def get_url(base_url: str, xml: str) -> str:

    url = '{}&{}'.format(base_url, parse.urlencode({'XML': xml}))
    return url

def set_alert(field_url: str,
              email_url: str,
              ca_path: str,
              request_field_xml: str,
              request_email_xml: str,
              tracking_number: str,
              email: str):

    """
    Set email alert for sender to receive future tracking activity
    """
    # Get current IP address to access Package Tracking "Fields" API
    IP_address = socket.gethostbyname(socket.gethostname())

    # Access to Package Tracking "Fields" API to collect MpSuffix and MpDate
    # These are required to access Tracking and Confirm by Email API
    request_field_url = get_url(base_url=field_url, 
                                xml=request_field_xml
                                .format(IP_address, tracking_number))
    
    get_mp_field = requests.get(request_field_url, verify=ca_path)
    result = get_mp_field.content.decode('utf-8')
    
    mp_suffix = re.search(r'[0-9]+(?=</MPSUFFIX>)', result)[0]
    mp_date = re.search(r'(?<=<MPDATE>)\S+\s\S+(?=</MPDATE>)', result)[0]
    
    # Access to Tracking and Confirm by Email API and set email alert
    request_email_url = get_url(base_url=email_url,
                                xml=request_email_xml
                                .format(tracking_number, mp_suffix, mp_date, email))
    response = requests.get(request_email_url, verify=ca_path)
    
    
def main():
    
    config = configparser.ConfigParser(interpolation=configparser.BasicInterpolation())
    config.read('config.txt')

    ca_path = config['DEFAULT']['ca_path']
    ups_headers = {'AccessLicenseNumber': config['UPS']['ups_key']}
    ups_url = config['UPS']['ups_url']
    
    field_url = config['USPS']['field_url']
    field_xml = config['USPS']['request_field']
    email_url = config['USPS']['email_url']
    email_xml = config['USPS']['request_email']

    email_account = config['EMAIL']['email_to_read']
    main_folder = config['EMAIL']['main_folder']
    sub_folder = config['EMAIL']['sub_folder']

    try:
        outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    except AttributeError:
        username = environ.get('USERNAME')
        file_loc = f"C:\\Users\\{username}\\AppData\\Local\\Temp\\gen_py"
        rmtree(file_loc, ignore_errors=True)
        outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')
        
    account = outlook.Folders[email_account]
    folder = account.Folders[main_folder]
    read_folder = folder.Folders[sub_folder]
    mail_items = read_folder.Items
    
    for mail in mail_items:
        
        if mail.Class == 43:
            mail_content = mail.Body
            find_tracking_number = re.findall(r'92\d{24}', mail_content)
            reply_all = mail.ReplyAll()
            
            if find_tracking_number != None and len(find_tracking_number) == 1:
                tracking_number = find_tracking_number[0]
                tracking_result = track_pic(url=ups_url,
                                            headers=ups_headers,
                                            ca_path=ca_path,
                                            tracking_number=tracking_number)
                
                if tracking_result == {}:
                    reply_all.HTMLBody = config['EMAIL']['invalid_tracking']
                elif len(tracking_result) == 4:
                    reply_all.HTMLBody = (config['EMAIL']['registered_tracking']
                                          .format(tracking_result['tracking_number'],
                                                  tracking_result['last_event'],
                                                  tracking_result['last_location'],
                                                  tracking_result['last_date_time'])
                                           + reply_all.HTMLBody
                    )
                    if mail.SenderEmailType == 'EX':
                        sender = mail.Sender.GetExchangeUser().PrimarySmtpAddress
                    else:
                        sender = mail.SenderEmailAddress
                        if '@ups.com' in sender.lower():
                            continue
                        else:
                            set_alert(field_url=field_url,
                                      email_url=email_url,
                                      ca_path=ca_path,
                                      request_field_xml=field_xml,
                                      request_email_xml=email_xml,
                                      tracking_number=tracking_number,
                                      email=sender)
                    
                elif len(tracking_result) == 3:
                    reply_all.HTMLBody = (config['EMAIL']['unregistered_tracking']
                                          .format(tracking_result['tracking_number'],
                                                  tracking_result['last_event'],
                                                  tracking_result['last_date_time'])
                                           + reply_all.HTMLBody
                    )

                reply_all.Save()
        
            else:
                continue
        else:
            continue
            

if __name__ == '__main__':
    main()