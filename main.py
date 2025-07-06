import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random

# modules
from oxford_db import import_oxford_db, save_word_to_oxford_db

OXFORD = 1
DEFINITION_COLUMN = 2
TC = "tc_00"
WORD_NOT_FOUND = "No exact match found for"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
}
MAX_TIME_WAITING = 2  # seconds
MIN_TIME_WAITING = 0.1  # seconds

oxford_db_dict = import_oxford_db()

df = pd.read_excel(TC + ".xlsx", engine="openpyxl", header=None)
for index, row in df.iterrows():
    # Extract the first column value
    first_col_value = row.iloc[0]
    print(f"Starting crawling word: {first_col_value}")

    # Check if word NOT in oxford_db_dict
    if first_col_value in oxford_db_dict:
        print(f"Word '{first_col_value}' exists in OXFORD DB")
        html_str = oxford_db_dict[first_col_value]
        df.at[index, DEFINITION_COLUMN] = html_str
        continue

    # find word on OXFORD
    url = f"https://www.oxfordlearnersdictionaries.com/search/english/?q={first_col_value}"

    response = requests.get(url, headers=HEADERS)

    if response.text.find(WORD_NOT_FOUND) != -1:
        print(f"Word '{first_col_value}' not found in OXFORD")
        continue

    # Parse HTML with BeautifulSoup
    soup = BeautifulSoup(response.text, "html.parser")

    oald = soup.find("div", class_="oald")
    # oald = soup.find("div", id="main-container")
    for tag in oald.find_all("div", id="ring-links-box"):
        tag.decompose()

    html_str = str(oald).replace("\n", "").replace("\r", "")

    # Add into column 2
    df.at[index, DEFINITION_COLUMN] = html_str
    # save word to oxford_db
    save_word_to_oxford_db(first_col_value, html_str)

    waiting_time_second = round(random.uniform(MIN_TIME_WAITING, MAX_TIME_WAITING), 10)
    time.sleep(waiting_time_second)


# save df into csv
df.to_csv(f"{TC}_testing.csv", index=False, header=None)
