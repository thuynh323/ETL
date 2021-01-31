import sys
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, InvalidArgumentException
from selenium.webdriver.chrome.options import Options
import sqlalchemy
import sqlite3
import pandas as pd
from datetime import datetime, timedelta

# Change with your path
ATABASE_LOCATION = 'sqlite:///brave_ad_tracker.sqlite'
DRIVER_PATH = r'C:\Users\Thanh Huynh\Documents\Projects\brave-ads\venv\chromedriver.exe'
BRAVE_PATH = r'C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe'
USER_DATA = 'C:\\Users\\Thanh Huynh\\AppData\\Local\\BraveSoftware\\Brave-Browser\\User Data'

def extract_data(user_data: str, brave_path: str, driver_path: str) -> dict:
    # Extract data from Brave browser
    # Turn off logs of decryption error and add user data
    option = Options()
    option.add_experimental_option('excludeSwitches', ['enable-logging'])
    option.add_argument('user-data-dir=' + user_data)
    option.binary_location = brave_path

    # Chedk if Brave is running
    # No exception, continue
    try:
        driver = webdriver.Chrome(executable_path= driver_path, options= option)
        driver.get('brave://rewards/')

        # Check if there were ads displayed
        # Scrape all ads
        ad_dict = {}
        try:
            button = driver.find_element_by_xpath("//*[contains(text(), '7-day Ads History')]")
            button.click()
            ad_table = driver.find_element_by_id('modal')
            boxes = ad_table.find_elements_by_xpath("//tbody/tr")
            for box in boxes:
                try:
                    ad_date = box.find_element_by_class_name('StyledDateText-sc-15yohet.fzlNYf').get_attribute('textContent')
                    ad_dict.update({ad_date: {}})
                    ad_dict[ad_date]['link'] = []
                    ad_dict[ad_date]['title'] = []
                    ad_dict[ad_date]['content'] = []
                    ad_dict[ad_date]['website'] = []
                    ad_dict[ad_date]['category'] = []
                except NoSuchElementException:
                    ad_link = box.find_element_by_class_name('StyledAdLink-sc-5m99w9.cbblyT').get_attribute('href')
                    ad_brand = box.find_element_by_class_name('StyledAdBrand-sc-cgfxyr.lavsD').get_attribute('textContent')
                    ad_info = box.find_elements_by_class_name('StyledAdInfo-sc-addoug.bzjRjV')
                    ad_description = ad_info[0].get_attribute('textContent')
                    ad_website = ad_info[1].get_attribute('textContent')
                    ad_category = box.find_element_by_class_name('StyledCategoryName-sc-10iobxf.eprdGa').get_attribute('textContent')
                    ad_dict[ad_date]['link'].append(ad_link)
                    ad_dict[ad_date]['title'].append(ad_brand)
                    ad_dict[ad_date]['content'].append(ad_description)
                    ad_dict[ad_date]['website'].append(ad_website)
                    ad_dict[ad_date]['category'].append(ad_category)
        except NoSuchElementException:
            print('No ads displayed in the last 7 days. Finishing execution')
        finally:
            driver.quit()
        return ad_dict
    # Brave is running. Terminate the program
    except InvalidArgumentException:
        print('Please close current Brave browser to proceed')
        sys.exit()
    
def transform_data(ad_dict: dict) -> pd.DataFrame:
    # Transform to a pandas dataframe
    ad_df = []
    for ad_date, ad in ad_dict.items():
        num_ads = len(list(ad.values())[0])
        ad['date'] = [ad_date]*num_ads
        df = pd.DataFrame.from_dict(ad)
        ad_df.append(df)
    ad_df = pd.concat(ad_df)
    ad_df = ad_df[['date', 'title', 'content', 'website', 'category', 'link']]
            
    # Select ads displayed yesterday only
    ad_df['date'] = ad_df['date'].apply(lambda x: datetime.strptime(x, '%m/%d/%Y'))
    yesterday = datetime.now() - timedelta(1)
    yesterday = datetime.strftime(yesterday, '%Y-%m-%d')
    to_store_df = ad_df[ad_df['date'] == yesterday]
    if to_store_df.empty:
        print('No ads displayed yesterday. Finishing execution')
        sys.exit()
    else:
        print('Data valid, proceed to Load stage')
        return to_store_df

def load_data(to_store_df: pd.DataFrame, database_location: str):
    
    # Create empty database
    engine = sqlalchemy.create_engine(database_location)
    conn = sqlite3.connect('brave_ad_tracker.sqlite')
    cursor = conn.cursor()
    sql_query = """
    CREATE TABLE IF NOT EXISTS brave_ad_tracker(
        ad_no INTEGER PRIMARY KEY AUTOINCREMENT,
        date VARCHAR(200),
        title VARCHAR(200),
        content VARCHAR(200),
        website VARCHAR(200),
        category VARCHAR(200),
        link VARCHAR(200)
    )
    """
    cursor.execute(sql_query)
    print('Opened database sucessfully')

    # Load in the data
    to_store_df.to_sql('brave_ad_tracker', engine, index= False, if_exists= 'append')
    conn.close()
    print('Loaded data sucessfully')

def main():
    extracted_data = extract_data(user_data = USER_DATA, brave_path = BRAVE_PATH, driver_path = DRIVER_PATH)
    transformed_data = transform_data(extracted_data)
    load_data(to_store_df = transformed_data, database_location = ATABASE_LOCATION)

if __name__ == '__main__':
    main()