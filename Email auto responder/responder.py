import warnings
import contextlib
import requests
from urllib3.exceptions import InsecureRequestWarning
import win32com.client as win32
from usps import USPSApi
import re

USER_ID = # Your USPS API username
EMAIL_ACCOUNT = # Your Outlook account

old_merge_environment_settings = requests.Session.merge_environment_settings

@contextlib.contextmanager
def no_ssl_verification():
    """
    Source: https://stackoverflow.com/a/15445989
    """
    opened_adapters = set()

    def merge_environment_settings(self, url, proxies, stream, verify, cert):
        # Verification happens only once per connection so we need to close
        # all the opened adapters once we're done. Otherwise, the effects of
        # verify=False persist beyond the end of this context manager.
        opened_adapters.add(self.get_adapter(url))
        settings = old_merge_environment_settings(self, url, proxies, stream, verify, cert)
        settings['verify'] = False
        return settings

    requests.Session.merge_environment_settings = merge_environment_settings
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', InsecureRequestWarning)
            yield
    finally:
        requests.Session.merge_environment_settings = old_merge_environment_settings
        for adapter in opened_adapters:
            try:
                adapter.close()
            except:
                pass

def track_pic(tracking_number: str) -> dict:
    """
    Returns a dictionary of tracking summary or tracking error
    """
    usps = USPSApi(USER_ID)
    response = usps.track(tracking_number).result.get('TrackResponse').get('TrackInfo')

    if 'TrackSummary' in response:
        result = response.get('TrackSummary')
    elif 'Error' in response:
        result = response.get('Error')
    else:
        pass
    return result

def tracking_summary(tracking_number: str, tracking_status: dict) -> str:
    
    pic_error = tracking_status.get('Number')
    pic_status = tracking_status.get('Event')
    pic_city = tracking_status.get('EventCity')
    pic_state = tracking_status.get('EventState')
    pic_zip = tracking_status.get('EventZIPCode')
    event_date = tracking_status.get('EventDate')
    event_time = tracking_status.get('EventTime')
    if pic_error == None:
        if pic_city != None:
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
                    &emsp;{pic_status}&emsp;{event_date} {event_time}<br>
                    USPS Entry Point:<br>
                    &emsp;{pic_city}, {pic_state} {pic_zip}</p>
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
                    &emsp;{pic_status}</p>
                    </font>
                </body>
                </html>
                """
    elif pic_error == '-2147219284':
        mail_body = f"""
                <html>
                <body>
                    <font face='Calibri' size='-0.5'>
                    <p>Hello,</p>
                    <p>Your request has been received and is being reviewed by our support department.
                    This package has currently been received by its destination processing facility.</p>
                    <p>Tracking number:<br>
                    &emsp;{tracking_number}<br></p>
                    </font>
                </body>
                </html>
                """
    else:
        mail_body = tracking_status.get('Description')
    return mail_body

def reply_mail(mail_items):
    for mail in mail_items:
        if '_MailItem' in str(type(mail)):
            mail_content = mail.Body
            find_tracking_number = re.search(r'92\d{24}', mail_content)
            if find_tracking_number != None:
                tracking_number = find_tracking_number.group(0)
                tracking_result = track_pic(tracking_number)
                reply_all = mail.ReplyAll()
                reply_all.HTMLBody = tracking_summary(tracking_number, tracking_result) + reply_all.HTMLBody
                reply_all.Save()
            else:
                continue
        else:
            continue

def main():
    outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    account = outlook.Folders[EMAIL_ACCOUNT]
    folder = account.Folders['Inbox']
    read_folder = folder.Folders['Test'] # Change with the appropriate folder

    with no_ssl_verification():
        reply_mail(read_folder.Items)

if __name__ == '__main__':
    main()
    
 
##########################--------------------------------------------------------################################

import requests
import win32com.client as win32
from datetime import datetime
import re

EMAIL_ACCOUNT = '' # Your Outlook email address
API_KEY = '' # Your Access License Number
HEADERS = {"AccessLicenseNumber": API_KEY}
API_VERSION = 'v1'
API_BASE_URL = f"https://onlinetools.ups.com/track/{API_VERSION}/details/"

def track_pic(tracking_number: str) -> dict:
    """
    Returns a dictionary of tracking summary or tracking error
    """
    r = requests.get(f'{API_BASE_URL}{tracking_number}', headers=HEADERS, verify=False)
    response = r.json().get('trackResponse').get('shipment')

    if response == None:
        tracking_summary = {}
    else:
        tracking_summary = response[0].get('package')[0].get('activity')[0]
    return tracking_summary

def tracking_details(tracking_number: str, tracking_summary: dict) -> str:
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
        location = tracking_summary.get('address')
        status = tracking_summary.get('status')

        last_city = location.get('city')
        last_state = location.get('stateProvince')
        last_zip = location.get('postalCode')
        last_status = status.get('description')
        last_date = datetime.strptime(status.get('date'), '%Y%m%d').date()
        last_time = datetime.strptime(status.get('date'), '%HH%MM%SS%').time()
    
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

def reply_mail(mail_items):
    for mail in mail_items:
        if '_MailItem' in str(type(mail)):
            mail_content = mail.Body
            find_tracking_number = re.search(r'92\d{24}', mail_content)
            reply_all = mail.ReplyAll()
            if find_tracking_number != None:
                tracking_number = find_tracking_number.group(0)
                tracking_result = track_pic(tracking_number)
                mail_body = tracking_details(tracking_number, tracking_result)
                reply_all.HTMLBody = mail_body + reply_all.HTMLBody
                reply_all.Save()
            else:
                continue
        else:
            continue

def main():
    outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
    account = outlook.Folders[EMAIL_ACCOUNT]
    folder = account.Folders['Inbox']
    read_folder = folder.Folders['Test'] # Change with the appropriate folder
    reply_mail(read_folder.Items)

if __name__ == '__main__':
    main()
