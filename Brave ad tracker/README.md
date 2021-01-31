# Brave ad tracker
<img src="https://github.com/thuynh323/ETL/blob/main/Brave%20ad%20tracker/photo/brave-logo.png" width="200" />

## Which ads were displayed while I was using Brave?
This project covers an ETL process based on my ads history from Brave rewards.

Since I could not find where Brave stores my ads history, I decided to build a pipeline to scrape these data from my own browser, then store them in an .sqlite database.
This process can be scheduled to run automatically via Windows Task Scheduler. In my case, I set up Task Scheduler to ask if I want to run this script whenever I log in my computer. Only data from ads showed the day before are appended to the existing database, so I just need to run this program once a day. Of course, there will be no data scraped if I do not log in my computer.

What am I going to do with these data? I might be able to generate some visualization or perform text mining. It would be nice if Brave provided specific time each ad was displayed, then I might try to find if there is a pattern of which ad they want me to watch and when.


- My ads history from Brave rewards

![alt text](https://github.com/thuynh323/ETL/blob/main/Brave%20ad%20tracker/photo/ad-history.png)
- Database

<img src="https://github.com/thuynh323/ETL/blob/main/Brave%20ad%20tracker/photo/database.png" width="1000" />

Motivated by: [Karolina Sowinska's ETL process built for Spotify data](https://www.youtube.com/watch?v=dvviIUKwH7o).

This mini series will show you how to retrieve data of songs you have listened to in the past 24 hour via Spotify's API. It also introduces Airflow as a scheduler to automate this task.

_Unfortunately, Windows does not support Airflow and I don't think I can scrape my browser in Ubuntu, so I decided to use Windows Task Scheduler_.
