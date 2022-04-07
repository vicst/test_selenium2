"""
A framework based on, used to control the browser.
"""

from selenium.webdriver.common.by import By
from traceback import print_stack
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
import app.utilities.custom_logger as cl
from app.utilities.custom_logger import screen_shot
import logging
import time
from retrying import retry


class SeleniumDriver(object):
    """
    Class used for controlling a browser.
    All the actions are logged using the custom_logger module.
    """

    log = cl.customLogger(logging.DEBUG)

    def __init__(self, driver):
        self.driver = driver

    def get_title(self):
        """
        Returns the title of the page
        """
        return self.driver.title

    def get_by_type(self, locatorType):
        """
        Gets the type of the locator for a given string
        Parameters
        ----------
        locatorType : str
        Type of locator can be: id, name, xpath, css, class, link

        Returns
        -------
        The locator type as defined by the BY class
        """
        locatorType = locatorType.lower()
        if locatorType == "id":
            return By.ID
        elif locatorType == "name":
            return By.NAME
        elif locatorType == "xpath":
            return By.XPATH
        elif locatorType == "css":
            return By.CSS_SELECTOR
        elif locatorType == "class":
            return By.CLASS_NAME
        elif locatorType == "link":
            return By.LINK_TEXT
        else:
            self.log.info("Locator type " + locatorType +
                          " not correct/supported")
        return False

    @retry(stop_max_attempt_number=3, wait_fixed=3)
    def get_element(self, locator, locatorType="id", take_screen_shot=True):
        """
        Gets a web element by its locator

        Parameters
        ----------
        locator : str
            The locator should be created using dev tools in browser. Check if is unique using search
        locatorType : str
             Type of locator can be: id, name, xpath, css, class, link

        Returns
        -------
            A selenium element object
        """
        element = None
        try:
            locatorType = locatorType.lower()
            byType = self.get_by_type(locatorType)
            element = self.driver.find_element(byType, locator)
            self.log.info("Element found with locator: " + locator +
                          " and  locatorType: " + locatorType)
        except:
            self.log.debug("Element not found with locator: " + locator +
                          " and  locatorType: " + locatorType)
            if take_screen_shot:
                screen_shot(self.driver, self.log)
        return element

    @retry(stop_max_attempt_number=3, wait_fixed=3)
    def get_elements(self, locator, locatorType="id"):
        """
        Gets a list of elements that are identified by their locator

        Parameters
        ----------
        locator : str
            The locator should be created using dev tools in browser. It can contain one or more elements
        locatorType : str
            Type of locator can be: id, name, xpath, css, class, link

        Returns
        -------
            A list of selenium elements

        """
        elements = None
        try:
            locatorType = locatorType.lower()
            byType = self.get_by_type(locatorType)
            elements = self.driver.find_elements(byType, locator)
            self.log.info("Element found with locator: " + locator +
                          " and  locatorType: " + locatorType)
        except:
            self.log.debug("Element not found with locator: " + locator +
                          " and  locatorType: " + locatorType)
            print("Element not found with locator: " + locator +
                          " and  locatorType: " + locatorType)
            screen_shot(self.driver, self.log)
        return elements

    #@retry(stop_max_attempt_number=3, wait_fixed=3)
    def element_click(self, locator, locatorType="id"):
        """
        Clicks on a element

        Parameters
        ----------
        locator : str
            The locator should be created using dev tools in browser. Check if is unique using search
        locatorType : str
            Type of locator can be: id, name, xpath, css, class, link
        """
        try:
            #element = self.get_element(locator, locatorType)
            element = self.wait_for_element(locator=locator, locator_type=locatorType,
                                            condition_text="element_to_be_clickable")
            element.click()
            self.log.info("Clicked on element with locator: " + locator +
                          " locatorType: " + locatorType)
        except Exception as e:
            self.log.debug("Cannot click on the element with locator: " + locator +
                          " locatorType: " + locatorType)
            screen_shot(self.driver, self.log)
            print(str(e))
            print_stack()

    @retry(stop_max_attempt_number=3, wait_fixed=3)
    def send_Keys(self, data, locator, locatorType="id"):
        """
        Writes in a input field.

        Parameters
        ----------
        data : str
            The text that should be writen in the input field
        locator : str
            The locator should be created using dev tools in browser. Check if is unique using search
        locatorType : str
            Type of locator can be: id, name, xpath, css, class, link

        """
        try:
            element = self.wait_for_element(locator=locator, locator_type=locatorType, refreshes=1)
            element.send_keys(data)

            # convert utf-8 return key
            special_keys = {'\ue006': "RETURN KEY", '\ue00c': "ESCAPE KEY", '\ue004': "TAB KEY",
                            "\ue003": "BACK SPACE"}
            if data in special_keys:
                data = special_keys[data]
            self.log.info("Sent " + str(data) + " on element with locator: " + locator +
                          " locatorType: " + locatorType)
            if data != "RETURN KEY":
                print("Send: " + str(data) + " was set")

            # FOR TESTING
            # if data != "RETURN KEY":
            #     print("TEXT: " + str(data) + " should be send")
            # self.log.info("___TEST send " + str(data) + " on element with locator: " + locator +
            #                " locatorType: " + locatorType)

        except Exception as e:
            self.log.debug("Cannot send data on the element with locator: " + locator +
                  " locatorType: " + locatorType)
            screen_shot(self.driver, self.log)
            print_stack()
            print(str(e))

    def is_element_present(self, locator, locatorType="id", take_screen_shot=True):
        """
        Checks the presence of an element
        """
        try:
            element = self.get_element(locator, locatorType, take_screen_shot)
            if element is not None:
                self.log.info("Element present with locator: " + locator +
                              " locatorType: " + locatorType)
                return True
            else:
                self.log.debug("Element not present with locator: " + locator +
                              " locatorType: " + locatorType)
                return False
        except:
            print("Element not found " + locator)
            if take_screen_shot:
                screen_shot(self.driver, self.log)
            return False


    def _check_load(self, element,  locator=None, locator_type="id"):
        """
        Just for testing
        :param element: (web element) the element(s) to be checked
        :return: element if it was found on page or False if element is not present(loaded)
        """
        try:
            locatorType = locator_type.lower()
            byType = self.get_by_type(locatorType)

            if type(element) == list:
                element = self.driver.find_elements(byType, locator)[0]
            else:
                element.find_element(locator_type, locator)
            return element
        except StaleElementReferenceException:
            return False

    def _retry_wait(self, element, locator=None, locator_type="id", find_time_out=3):
        "Just for testing"
        start_time = time.time()
        while time.time() < start_time + find_time_out:
            if self._check_load(element, locator, locator_type):
                return element
            # else:
            #     new_el = self._check_load(element, locator, locator_type)
        screen_shot(self.driver, self.log)
        return None

    #@retry(stop_max_attempt_number=3, wait_fixed=3)
    def wait_for_element(self, *args, locator=None,
                         locator_type="id",
                         timeout: int = 7,
                         poll_frequency=0.5,
                         condition_text="presence_of_element_located",
                         refreshes=0,
                         check_page_load=False,
                         take_screen_shot=True):
        """Checks for certain conditions to be fulfilled. Depending on the condition it will return a web element,
        some web elements or a boolean. Check conditions link for more details:
        https://seleniumhq.github.io/selenium/docs/api/py/webdriver_support/selenium.webdriver.support.expected_conditions.html

        Parameters
        ---------
        args: str
        Used to create expected condition check the above link to get the args. Locator and locator type
            are defined separately, not in the args
        locator: str
            The locator for the element(s) to be checked
        locator_type: str
            The type of the locator
        timeout: str
            The time to wait for a condition to be fulfilled
        poll_frequency: str
            The time between retries for condition check. The retry is done until timeout
        condition_text: str
            The text of the condition. Please check the link above for text
        refreshes: int
            Refresh the page and try again if it didn't mach the condition

        Returns
        -------
            Element(s) or boolean
        """


        if locator:
            byType = self.get_by_type(locator_type)
            if len(args) != 0:
                expression = "EC.{condition}(('{locator_type}', '{locator}'),{other_args})".format(condition=condition_text,
                                                                                                   locator_type=byType,
                                                                                                   locator=locator,
                                                                                                   other_args=args[0])
            else:
                expression = "EC.{condition}(('{locator_type}', '{locator}'))".format(
                    condition=condition_text,
                    locator_type=byType,
                    locator=locator)
        else:
            expression = "EC." + condition_text + str(args)

        condition = eval(expression)

        wait = WebDriverWait(self.driver, timeout, poll_frequency,
                             ignored_exceptions=[NoSuchElementException,
                                                 ElementNotSelectableException])
        wait_result = None
        try:
            wait_result = wait.until(condition)
            # Refresh the page and try again if it didn't mach the condition
        except TimeoutException:
            if take_screen_shot:
                screen_shot(self.driver, self.log)
            print("Element timeout: " + locator + " " + "Page should refresh " + str(refreshes) + " times")
            if refreshes != 0:
                for refresh_count in range(refreshes):
                    if not wait_result:
                        time.sleep(10)
                        self.driver.refresh()
                        print("Page refreshed")
                        wait_result = wait.until(condition)
                        if wait_result is not None:
                            return wait_result
                        print("Element timeout after refresh")
            else:
                self.log.debug("Cannot wait for element with locator: " + locator +
                               " locatorType: " + locator_type)

        except Exception as e:
            self.log.debug("Cannot wait for element with locator: " + locator +
                  " locatorType: " + locator_type)
            screen_shot(self.driver, self.log)
            print(str(e))

        # If check_page_load is True, it tries to find the element on the page and retries if it was not found
        if check_page_load:
            self._retry_wait(wait_result, locator, locator_type)
        return wait_result

    def get_text(self, locator=None,
                         locator_type="id",
                         timeout: int = 7,
                         poll_frequency=0.5,
                         condition_text="presence_of_element_located",
                         refreshes=0,
                         check_page_load=False,
                         take_screen_shot=True):
        element = self.wait_for_element(locator=locator,
                              locator_type=locator_type,
                              timeout=timeout,
                              refreshes=refreshes,
                              take_screen_shot=take_screen_shot)
        element_text = ''
        if element:
            try:
                element_text =element.text
            except Exception as e:
                self.log.debug("Cannot get text for element with locator: " + locator +
                               " locatorType: " + locator_type + " exception: " + repr(e))
        return element_text