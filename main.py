# Ground News Top News Stories
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options # Changes default selenium options
from selenium.webdriver.chrome.service import Service
from datetime import datetime
import os
import sys
import openpyxl # Explicit so it is recognized by pyinstaller. Could also add to main.spec

# Get the path of the executable we are going to create
application_path = os.path.dirname(sys.executable)

now = datetime.now()
month_day_year = now.strftime(r'%m%d%Y') # MMDDYYYY

url = "https://ground.news/"
path = r"C:\Users\thai\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe"


# headless-mode (doesn't open tab)
options = Options()
options.add_argument('--headless')

service = Service(executable_path=path)

driver = webdriver.Chrome(service=service, options=options)
driver.get(url)


# Container XPath locations
containers = driver.find_elements(by="xpath", value="//div/div[contains(@class, 'group')]/div[@class='relative']/a")

headlines = []
coverages = []
links = []

# click through home button
close_button = driver.find_element(by="xpath", value="//button[contains(text(), 'Ground News homepage')]").click()

for container in containers:
    # Headline HTML XPath
    headline = container.find_element(by="xpath", value="./div/div/h4[text()]").text

    # Coverage HTML XPath
    coverage = container.find_elements(by="xpath", value="./div/div/div/span")
    
    biases = []
    for bias in coverage:
        biases.append(bias.text)
    
    # Link HTML XPath
    link = container.find_element(by="xpath", value=".").get_attribute('href')

    headlines.append(headline)
    coverages.append(biases)
    links.append(link)


# Creating a dictionary to turn into a DataFrame
dict = {'headline': headlines, 'coverage': coverages, 'link': links}
df = pd.DataFrame(dict)
# Only keeps rows where condition is True (non-empty)
df = df[df['headline'] != '']

file_name = f'headlines-{month_day_year}.xlsx'
# Creating a path by joining two path variables to avoid \/
final_path = os.path.join(application_path , file_name)

try:
    df.to_excel(final_path, index=False)
except PermissionError:
    print('CLOSE EXCEL DUDE')

driver.quit()