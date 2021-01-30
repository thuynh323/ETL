# Brave ad tracker
<img src="https://github.com/thuynh323/ETL/blob/main/Brave%20ad%20tracker/photo/brave-logo.png" width="200" />

## Which ads were displayed while I was using Brave?
This project covers an ETL process based on my ads history from Brave rewards.

Since I could not find where Brave stores my personal ads history, I decided to build a pipeline to scrape these data from my own browser, then store them in an .sqlite database.
This process will be scheduled to run daily using Airflow.

Motivated by: [Karolina Sowinska's ETL process built for Spotify data](https://www.youtube.com/watch?v=dvviIUKwH7o)
- My ads history from Brave rewards

![alt text](https://github.com/thuynh323/ETL/blob/main/Brave%20ad%20tracker/photo/ad-history.png)
- Database

<img src="https://github.com/thuynh323/ETL/blob/main/Brave%20ad%20tracker/photo/database.png" width="1000" />
