from selenium import webdriver
import time
import shutil
from selenium.webdriver.common.by import By

# this script downloads the Excel in DOWNLOADS !!!

# using selenium's chromedriver to apply a web scraping technique
# that finds the button i'm looking for (Genereaza Excel)
# and clicks it
# which will automatically download a table with the latest data about energy consumption and production
webdriver_path = "C:/path/to/chromedriver.exe"
excel_file_path = "C:/Users/paula/source/repos/data-extraction-to-html/resources/Grafic_SEN.xlsx" # doesnt work -> wanted to save the Excel in a resources directory from where i could retrieve using script_ExtractToHTML but the file kept getting corrupted during the process # idk still need it for the code to work !!!
driver = webdriver.Chrome(executable_path=webdriver_path)

# getting the web page
url = "https://www.transelectrica.ro/widget/web/tel/sen-grafic/-/SENGrafic_WAR_SENGraficportlet"
driver.get(url)

# looking for the SEN_search button (= Genereaza Excel)
# after the click -> gave it 10s time to download just in case but the web page will close after 10s
button = driver.find_element(By.XPATH, '//span[contains(@onclick, "SEN_search(true)")]')
button.click()
time.sleep(10)

# downloaded_file_path = r"C:\Users\paula\source\repos\data-extraction-to-html\resources\Grafic-SEN.xlsx"
# shutil.move(downloaded_file_path, excel_file_path)

# needed to close the browser window which will open briefly after i run the script
driver.quit()