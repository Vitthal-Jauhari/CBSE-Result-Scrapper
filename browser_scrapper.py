from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.remote_connection import LOGGER
from lib2to3.pgen2 import driver
import logging
import time


# Suppress Selenium console warning to suppress warnings and declutter console output
LOGGER.setLevel(logging.WARNING)


def CalcPercent(marks):
    try:
        best5 = marks[:]                        # Slicing to make a new copy rather than 2 lists with same memory address
        best5.remove(min(marks))                    
        print("best5: ", best5)
        best5 = (sum(best5)/500)*100
        best5 = round(best5, 2)                 # Rounding off to nearest 2 decimal places             

        core5 = marks[:]                        # Slicing to make a new copy rather than 2 lists with same memory address 
        core5 = core5[0:5]                  
        print("core5: ", core5)
        core5 = (sum(core5)/500)*100
        core5 = round(core5, 2)                 # Rounding off to nearest 2 decimal places

        return (best5, core5)
    except Exception as e:
        print("Error while calculating percentage: ", e)


class Browser:
    browser , service = None, None

    def __init__ (self, driver: str):
        self.service = Service(driver)
        self.browser = webdriver.Chrome(service=self.service)


    def open_page(self, url: str):
        self.browser.get(url)

    def close_page(self):
        self.browser.close()

    def add_input(self, by: By, value: str, text: str):
        field = self.browser.find_element(by=by, value=value)
        field.send_keys(text)
        # time.sleep(1)

    def click_button(self, by: By, value: str):
        button = self.browser.find_element(by=by, value=value)
        button.click()

    def login_cbse(self, roll_number:int, admit_card_ID:str):
        self.add_input(by=By.NAME, value='regno', text=roll_number)
        self.add_input(by=By.NAME, value='sch', text='70968')
        self.add_input(by=By.NAME, value='admid', text=admit_card_ID)
        self.click_button(by=By.NAME, value='B2')


    def read_table_data(self):
        try:
            table = self.browser.find_element(By.XPATH, '//html/body/div[1]/div/center/table/tbody')
            data = []  # All the text on the result page
            for tr in table.find_elements(By.XPATH, './/tr'):
                row = [item.text for item in tr.find_elements(By.XPATH, './/td')]
                data.append(row)
            return data
        except Exception as e:
            print("Error while reading table data: ", e)


    def Format_data(self, data):
        marks = []              # Will contain 6 integers in a list for each subject
        try:
            for k in data:
                try:
                    if k[0] in ['301', '041', '042', '043', '083', '836']:              # Matching against subject code
                        marks.append( int(k[-2]))
                except:
                    continue
        except Exception as e:
            print("Error during formatting: ", e)
        return marks
    

def fetch_details(name, rollno, admitcardID):
    # Specifying browser for operations ['Chromedriver']
    browser = Browser('C:\\Users\\vitth\\AppData\\Local\\Programs\\Python\\Python310\\Codes\\Selenium Projects\\drivers\\chromedriver.exe')
    # time.sleep(0.5)

    browser.open_page('https://cbseresults.nic.in/class_xiith_a_2024/ClassTwelfth_c_2024.htm')
    # time.sleep(1)

    browser.login_cbse(rollno, admitcardID)
    # time.sleep(3)

    scrapped_data = browser.read_table_data()
    formatted_data = browser.Format_data(data=scrapped_data)
    print(formatted_data)

    i,j = CalcPercent(formatted_data)
    # time.sleep(5)
    final_data = [name , rollno, admitcardID, i, j]
    # browser.open_page('https://cbseresults.nic.in/class_xiith_a_2024/ClassTwelfth_c_2024.htm')
    return final_data


if __name__ == '__main__':
    # This script is not supposed to be called directly
    # Sample data for testing purpose only
    # fetch_details(name='Vitthal Jauahri', rollno='23718733', admitcardID='VA737072')
    pass