__author__ = 'Kan!skA'

import unittest
import csv
import time, os
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

class TTS(unittest.TestCase):
    driver = None

    @classmethod
    def setUpClass(cls):
        cls.driver = webdriver.Firefox()
        time.sleep(11)
        cls.driver.maximize_window()
        cls.driver.get("http://intranet.daiict.ac.in:8085/DonlabTTS/")

    def test_TTS(self):
        # Change Your Set Number Here
        set = "Set 2"

        rows = dict()
        data_file = open(set + '/Paragraph/Paragraphs.csv', "r", encoding='utf-8')
        reader = csv.reader(data_file)
        lineNum = 1
        for row in reader:
            rows[lineNum] = row
            lineNum += 1
        for inputLineNum, inputData in rows.items():
            # Click on Clear Button
            self.driver.find_element_by_xpath(".//*[@id='AutoNumber1']/tbody/tr[4]/td/input[1]").click()
            time.sleep(1)
            # Enter Input in the Text Area
            self.driver.find_element_by_xpath(".//*[@id='AutoNumber1']/tbody/tr[7]/td[2]/p/textarea").send_keys(inputData)
            time.sleep(2)
            # Click on Synthesize Button
            self.driver.find_element_by_xpath(".//*[@id='AutoNumber1']/tbody/tr[4]/td/input[2]").click()
            time.sleep(5)
            # Right-Click on Output Link
            download = self.driver.find_element_by_xpath(".//*[@id='jsn-page']/center[1]/a")
            ActionChains(self.driver).context_click(download).key_down(Keys.ARROW_DOWN).perform()
            time.sleep(2)
            shell = win32com.client.Dispatch("WScript.Shell")
            shell.SendKeys('k')
            time.sleep(5)
            # Path for Keeping Output
            str1 = str('\Audio\\' + set + '\Paragraph\P-' + set + '-')
            shell.SendKeys(os.getcwd() + str1 + str(inputLineNum))
            time.sleep(5)
            shell.SendKeys('{ENTER}')
            time.sleep(2)
            # Click on Back
            self.driver.find_element_by_xpath(".//*[@id='jsn-page']/center[2]/a").click()

    @classmethod
    def tearDownClass(cls):
        cls.driver.quit()

if __name__ == '__main__':
    unittest.main()