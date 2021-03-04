import requests
import win32com.client as win32
import configparser
import json
import re
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
    
def main():
    
    config = configparser.ConfigParser(interpolation=configparser.BasicInterpolation())
    config.read('config.txt')

    email_account = config['EMAIL']['email_to_read']
    ups_key = config['DEFAULT']['ups_key']
    headers = {'AccessLicenseNumber': ups_key}
    ca_path = config['DEFAULT']['ca_path']
    url = config['DEFAULT']['base_url']
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
                tracking_result = track_pic(url=url,
                                            headers=headers,
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