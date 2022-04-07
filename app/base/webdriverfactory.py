"""
@package base

|WebDriver Factory class implementation
|It creates a webdriver instance based on browser configurations

Example:
    wdf = WebDriverFactory(browser)
    wdf.getWebDriverInstance()
"""
from selenium import webdriver
from app.utilities.custom_logger import screen_shot
import app.utilities.custom_logger as cl
import logging
import sys
import os
import time


class WebDriverFactory():
    """Used to create and configure the webdriver
    """

    log = cl.customLogger(logging.DEBUG)

    def __init__(self, browser, base_url):
        """
        Inits WebDriverFactory class
        """
        self.browser = browser
        self.base_url = base_url

    def getWebDriverInstance(self):
        """
        Set chrome driver and iexplorer environment based on OS

        chromedriver = "C:/.../chromedriver.exe"
        os.environ["webdriver.chrome.driver"] = chromedriver
        self.driver = webdriver.Chrome(chromedriver)

        PREFERRED: Set the path on the machine where browser will be executed

        Returns:
            'WebDriver Instance'
        """
        if self.browser == "iexplorer":
            # Set ie driver
            driver = webdriver.Ie()
        elif self.browser == "firefox":
            driver = webdriver.Firefox(executable_path=os.path.join(os.path.split(sys.argv[0])[0], "geckodriver.exe"))
        elif self.browser == "chrome":
            # Set chrome driver
            options = webdriver.ChromeOptions()
            #options.add_argument("headless")
            options.add_argument(r"user-data-dir=D:\Users\eddiehamilton\AppData\Local\Google\Chrome\User Data\Default")
            driver = webdriver.Chrome(os.path.join(os.path.split(sys.argv[0])[0], "chromedriver.exe"),
                                      chrome_options=options)
        # Setting Driver Implicit Time out for An Element
        #driver.implicitly_wait(15)
        # Setting load timeout
        driver.set_page_load_timeout(60)
        # Maximize the window

        #driver.minimize_window()
        # Loading browser with App URL. Try 3 times
        for i in range(3):
            try:
                driver.get(self.base_url)
                #driver.set_window_size(width=1980, height=1050)
                driver.maximize_window()
                # driver.set_window_position(0, 0)
                print("Page opened " + self.base_url + "  " + str(driver.get_window_size()))
                self.log.info("Page opened " + str(driver.get_window_size()) + ' ' + self.base_url)
                #driver.minimize_window()
                return driver
            except:
                self.log.error(self.base_url + " could not be loaded!")
                print(self.base_url + " could not be loaded!")
                time.sleep(10)
                driver.refresh()
                screen_shot(driver, self.log)


print(os.path.join(os.path.split(sys.argv[0])[0], "chromedriver.exe"))

if __name__ == "__main__":
    print(os.path.join(os.path.split(sys.argv[0])[0], "chromedriver.exe"))