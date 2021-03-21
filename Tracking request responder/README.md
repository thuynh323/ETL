# Respond to tracking request emails
This script automatically drafts email responses to tracking requests and sets email notifications for the senders.
## Overview
My team has received a large number of tracking request emails everyday, and this number can increase rapidly during peak seasons. This project was created to help save time of manually responding to each email.

Response exmaple:
<details>
  <summary>USPS registered tracking number</summary>
  
>Hello,
>
>Your request has been received and is being reviewed by our support department. While we investigate this package, we have set up an email alert with USPS for you to receive updates until the package is delivered.
>
>Tracking number:<br>
>&emsp;&emsp;&emsp;&emsp;92612902338293553000561745<br>
>Current package status:<br>
>&emsp;&emsp;&emsp;&emsp;Departed UPS Facility<br>
>Current location:<br>
>&emsp;&emsp;&emsp;&emsp;Urbancrest, OH 43123<br>
>Date, time :<br>
>&emsp;&emsp;&emsp;&emsp;01-12-2021 3:32
</details>

<details>
  <summary>USPS unregistered tracking number</summary>
  
>Hello,
>
>Your request has been received and is being reviewed by our support department. Please see the latest tracking event below.
>
>Tracking number:<br>
>&emsp;&emsp;&emsp;&emsp;92612902338293553000561745<br>
>Current package status:<br>
>&emsp;&emsp;&emsp;&emsp;Order information received<br>
>Date, time :<br>
>&emsp;&emsp;&emsp;&emsp;01-12-2021 3:32
</details>

<details>
  <summary>Invalid tracking number</summary>
  
>Hello,
>
>Unfortunately, we are unable to locate this package in our system.
</details>


These are requests for [UPS Mail Innovations Team](https://www.ups.com/us/en/services/shipping/mail-innovations.page) packages. Due to the task requirements, we need to provide the latest tracking events from UPS data and set up USPS email notifications for the customers.
## Configuration
The program is driven by a configuration file (config.ini). All sections are required, and one might need to change values in some sections to have the script run properly.
1. **UPS API access key**: See instructions [here](https://www.ups.com/upsdeveloperkit?loc=en_US)
2. **USPS user id**: See instructions [here](https://www.usps.com/business/web-tools-apis/)
3. **Access to USPS Track and Confirm by Email API**: A Mailer ID is required to get this access. See more [here](https://www.usps.com/business/web-tools-apis/track-and-confirm-api_files/track-and-confirm-api.htm#_Toc41911520)
4. **Certificate Authority .pem file**: May be required to avoid SSL errors while running the script from a coporate computer
