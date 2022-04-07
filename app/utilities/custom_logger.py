import inspect
import logging
from datetime import date, datetime
import os
import sys


def customLogger(logLevel=logging.DEBUG, logs_dir_name="iController_Logs"):
    """
    Creates the logger object
    Parameters
    ----------
    logLevel
    logs_dir

    Returns
    -------

    """
    logs_dir = os.path.join((os.path.split(sys.argv[0])[0]), logs_dir_name)
    #print("Logs will be saved to: " + logs_dir)
    # Create log dir if doesn't exists
    if not os.path.exists(logs_dir):
        os.mkdir(logs_dir)

    log_file = os.path.join(logs_dir, "automation_" + str(date.today()) + ".log")

    # Gets the name of the class / method from where this method is called
    loggerName = inspect.stack()[1][3]
    logger = logging.getLogger(loggerName)
    # By default, log all messages
    logger.setLevel(logging.DEBUG)

    fileHandler = logging.FileHandler(log_file, mode='a')
    fileHandler.setLevel(logLevel)

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s: %(message)s',
                    datefmt='%m/%d/%Y %I:%M:%S %p')
    fileHandler.setFormatter(formatter)
    logger.addHandler(fileHandler)
    return logger


def screen_shot(driver, log, screenshots_directory_name="Screenshoots"):
    """
    Takes screenshot of the current open web page

    :param driver: webdriver instance
    :param log: logger instance
    """
    screenshots_directory = os.path.join((os.path.split(sys.argv[0])[0]), screenshots_directory_name)
    file_name = inspect.stack()[1][3] + datetime.now().strftime("_%Ih_%Mm_%Ss") + ".png"
    if not os.path.exists(screenshots_directory):
        os.mkdir(screenshots_directory)
    screenshots_directory = os.path.join(screenshots_directory, "screenshoots_" + str(date.today()))
    relative_file_name = os.path.join(screenshots_directory, file_name)
    current_directory = os.path.dirname(sys.argv[0])
    destination_file = os.path.join(current_directory, relative_file_name)
    destination_directory = os.path.join(current_directory, screenshots_directory)
    try:
        if not os.path.exists(destination_directory):
            os.makedirs(destination_directory)
        driver.save_screenshot(destination_file)
        print("Screenshot save to directory: " + destination_file)
        log.info("Screenshot save to directory: " + destination_file)
    except Exception as e:
        log.error("### Exception Occurred when taking screenshot")
        print(e)

