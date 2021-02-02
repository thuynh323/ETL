import warnings
import contextlib
import requests
from urllib3.exceptions import InsecureRequestWarning
import win32com.client as win32
from usps import USPSApi
import re
import json

USER_ID = '' # Your USPS API username
EMAIL_ACCOUNT = '' # Your Outlook email address
FIRST_FOLDER = '' # Your folder/ first folder

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
                <style>
                    .tab {
                        tab-size: 8;
                    }
                </style>
                <body>
                    <font face='Calibri' size='-0.5'>
                    <p>Hello,</p>
                    <p>Your request has been received and is being reviewed by our support department. 
                    While we investigate this package, we have set up an email alert with USPS for you
                    to receive updates until the package is delivered.</p>
                    Tracking number:<br>
                    <pre class='tab'>   {tracking_number}<br></pre>
                    Current package status:<br>
                    <pre class='tab'>   {pic_status}&emsp;{event_date} {event_time}<br></pre>
                    USPS Entry Point:<br>
                    <pre class='tab'>   {pic_city}, {pic_state} {pic_zip}</pre></p>
                    </font>
                </body>
                </html>
                """
        else:
            mail_body = f"""
                <html>
                <style>
                    .tab {
                        tab-size: 8;
                    }
                </style>
                <body>
                    <font face='Calibri' size='-0.5'>
                    <p>Hello,</p>
                    <p>Your request has been received and is being reviewed by our support department. 
                    While we investigate this package, we have set up an email alert with USPS for you
                    to receive updates until the package is delivered.</p>
                    Tracking number:<br>
                    <pre class='tab'>   {tracking_number}<br></pre>
                    Current package status:<br>
                    <pre class='tab'>   {pic_status}</pre></p>
                    </font>
                </body>
                </html>
                """
    if pic_error == '-2147219284':
        mail_body = f"""
                <html>
                <style>
                    .tab {
                        tab-size: 8;
                    }
                </style>
                <body>
                    <font face='Calibri' size='-0.5'>
                    <p>Hello,</p>
                    <p>Your request has been received and is being reviewed by our support department. 
                    While we investigate this package, we have set up an email alert with USPS for you
                    to receive updates until the package is delivered.</p>
                    Tracking number:<br>
                    <pre class='tab'>   {tracking_number}<br></pre>
                    Current package status:<br>
                    <pre class='tab'>   {pic_status}</pre></p>
                    </font>
                </body>
                </html>
                """
    return mail_body

outlook = win32.gencache.EnsureDispatch('Outlook.Application').GetNamespace('MAPI')
account = outlook.Folders[EMAIL_ACCOUNT]
folder = account.Folders[FIRST_FOLDER]
read_folder = folder.Folders['Test']

with no_ssl_verification():

    for mail in read_folder.Items:
        if '_MailItem' in str(type(mail)):
            mail_content = mail.Body
            find_tracking_number = re.findall(r'92\d{24}', mail_content)
            if len(find_tracking_number) == 1:
                tracking_number = find_tracking_number[0]
                tracking_result = track_pic(tracking_number)
                reply_all = mail.ReplyAll()
                if 'Number' in tracking_result:
                    reply_all.HTMLBody = f"Hello,\n\n{tracking_result.get('Description')}.\n\nHave a great day!" + reply_all.HTMLBody
                else:
                    reply_all.HTMLBody = tracking_summary(tracking_number, tracking_result) + reply_all.HTMLBody
                reply_all.Save()
            else:
                continue
        else:
            continue
