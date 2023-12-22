from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import Chrome
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import pandas as pd 
import os
import time
df =pd.read_excel("input_folder\Top_RankKeywords.xlsx")
options = Options()
path = r"D:\BotPy\input_folder"
options.add_experimental_option("prefs",{"download.default_directory": path})
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.get('https://app.neilpatel.com/en/login')  

def get_latest_downloaded_file(directory):
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.csv')]  # Assuming it's a CSV file
    # if files:
    latest_file = max(files, key=lambda x: os.path.getmtime(os.path.join(directory, x)))
    # return os.path.join(directory, latest_file)
    return latest_file

#enter your email password 
email =""
password =""
driver.find_element(By.NAME, "email").send_keys(email)
driver.find_element(By.NAME, "password").send_keys(password)
login_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="login-button"]'))
)

# Click the button
login_button.click()
enterU = input("press y once inside KAMI")

if enterU == "y" or enterU =="Y":
    driver.get('https://app.neilpatel.com/en/traffic_analyzer/keywords')
    startingRow =input("enter the row from where you would like to start")
    startingRow = int(startingRow)
    batchSize =input("kami enter number of rows you want to process")
    
    batchSize = int(batchSize)
    print (batchSize)
    rangeval = startingRow+batchSize
    print (rangeval)
    for i in range(startingRow,rangeval):
        print(i)
        search_Value =df['DOMAINS'][i]
        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[data-testid="search-input"]'))
        )

        # Send keys to the search input field
        search_input.send_keys(search_Value)
        time.sleep(0.2)
        search_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="search-button"]'))
        )

        # Click the Search button
        try:
            search_button.click()
            export_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[1]/div/button'))
            )

            # Click the Export to CSV button
            export_button.click()

            dropdown_option = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div[2]/div[2]/div/div[1]/div[1]/div/div[2]/div/div[1]/div[2]/div[1]'))
            )

            # Click the dropdown option
            dropdown_option.click()
            time.sleep(50)
            fileName = get_latest_downloaded_file(path)
            current_file_name = fileName
            new_file_name  = "new_filename{}.csv".format(i)

            # Assuming the path to the directory containing the file
            directory_path = "D:\\BotPy\\input_folder"

            # Full path to the current and new files
            current_file_path = os.path.join(directory_path, current_file_name)
            new_file_path = os.path.join(directory_path, new_file_name)

            # Rename the file
            os.rename(current_file_path, new_file_path)
            outputExcel =new_file_path
            print(outputExcel)
            dfkeyword = pd.read_csv(outputExcel)
            dfkeyword['Keywords'] = dfkeyword['Keywords'].astype(str)

            combined_keywords = ', '.join(dfkeyword['Keywords'])
        except:
            combined_keywords = ""
        df['Top 1,000 keywords'][i]= combined_keywords
        search_input.clear()
df.to_excel("input_folder\Top_RankKeywords.xlsx")
