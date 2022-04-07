"""
The automation for the main page.
"""
from itertools import islice

from app.base.selenium_driver import SeleniumDriver
from app.utilities.create_report import CreateReport
import os
from app.configFiles.rules import LabelRules
from random import choice
import logging
import app.utilities.custom_logger as cl
from app.utilities.nice_tools import send_error_email
from selenium.webdriver.common.keys import Keys
from sys import exit as exit_app
import sys
import time
from app.utilities.custom_logger import screen_shot


class WebInterations(SeleniumDriver):
    """
        Class used for controlling a iController main page using
        SeleniumDriver methods.

        Attributes
        ----------
            _subject_locator : str
                locator for subject
            _subject_from_popup_locator : str
                locator for subject_from_popup
            _sender_locator : str
                locator for sender
            _date_locator : str
                locator for date
            _label_locator : str
                locator for label
            _email_body_locator : str
                locator for email body
            _label_field_locator : str
                locator for  label field

            driver :  webdriver object
                Should be constructed using getWebDriverInstance method from
                WebDriverFactory class
            label_rules_obj : str
                Excel file that defines the rules
            main_page_url : str
                The url of iController app
            country : str
                The country used for running the automation. It can be
                Belgium or Luxemburg
            sender_words : list
                Words used for defining current email body
            next_email_words : list
                Word used for skipping the first occurrence of the
                 sender word

            last_mail_date : str
                Date of the last mail taken from last report
            stop : Boolean
                Set to True if current email date <= last mail date

    """

    # LOCATORS
    # Check if page is opened (new message locator)
    _first_checkbox_locator = "#messages-list > tbody > tr:nth-child(1)"

    # Mail details
    _subject_locator = "#messages-list > tbody > tr > td.column-subject > a"
    _subject_from_popup_locator = "body > div > form > h2"
    _sender_locator = "#messages-list > tbody > tr > td.column-from > span"
    _sender_from_popup_locator = "body > div.main-content > form > div.from"
    _receiver_from_popup_locator = "body > div > form > div.contact > div > span"
    _date_locator = "#messages-list > tbody > tr> td.column-received-at.sorting_1"
    _label_locator = "#messages-list > tbody > tr > td.column-3"

    # Mail body
    _email_body_locator = "body > div > form > div.content"
    _label_field_locator = "s2id_autogen4"
    _clients_field_locator = "#s2id_autogen2"
    _label_body_locator = "#s2id_autogen3 > ul > li > div"
    _label_dropdown_element = "#select2-drop > ul > li:nth-child(1)"
    _done_button_locator = "body > div > form > div.actions > button.done.call-to-action"

    _client_dropdown_label_locator = "#select2-drop > ul > li > ul > li:nth-child(1)"

    # Forward window
    _forward_field_locator = "#s2id_autogen4"
    _forward_button_locator = "body > div.main-content > form > div.actions > button.send.call-to-action"
    _forward_accepted_locator = "#s2id_autogen3 > ul > li.select2-search-choice"

    log = cl.customLogger(logging.DEBUG)
    error_receiver = "victor.stanescu@leaseplan.com"

    def __init__(self, driver, label_rules_obj, clients_rules_obj):
        super().__init__(driver)
        self.driver = driver

        self.label_rules_obj = label_rules_obj
        self.clients_rules_obj = clients_rules_obj

        #self.main_page_url = main_page_url

        self.sender_words = [
            "From:", "From :",
            "Van:", "Van :",
            "De:", "De :",
            "To:", "To :",
            "Aan:", "Aan :",
            "A:", "A :"
        ]

        # If those words are find, the mail anterior mail will be considered
        self.next_email_words = ["la demande ci-dessous",
                                 "please see below",
                                 "zie de onderstaande",
                                 "zie de onderstaande bericht",
                                 "FYI"]

        # The labels will be set from last mail from the last report to the mail with this date and subject
        self.last_mail_date = None
        self.stop = False

    def get_emails_details(self):
        """
        1. Get a dictionary that contains the subjects, urls, senders, dates, and labels from the main page
        2. Unpack the dictionary into lists emails_subjects, emails_urls, mails_senders, mails_dates, mail_labels
        3. Create empty list that will be filled with the info that will be writen in the report: current_emails,
        labels_list, subject_list, status_list, labels_check_text_list, url_list
        4. Iterate simultaneously the list form main page dictionary and:
            a) check if the email has the same date as the last email from previous report. Break if it does
            b) check if the mail has the label already set, skip to next mail. Add info from main page lists to report
            lists and continue to the next email
            c) open current email pop-up and use open_body to get label, current_email, labels_check_text, subject_text.
            open_body method also sets the labels. Add information extracted from current email pop-up to the reports
            lists
            d) the lists from main page dictionary are replaced by the reports lists

        Returns the dictionary that contains the mails details. It checks if labels were already been set and
        sets the labels for the rest of the emails.

        Returns
        -------
        A dictionary that contains the mails details
            {"subject" : emails_subjects,
            "sender" : 	mails_senders
            "date" : mails_dates
            "label" : mail_labels
            "status" : status_list
            "email_body" : current_emails
            "check" : labels_check_text_list
            "url" : url_list
            "client" : client_list
            "clients_check_text" : clients_check_list
            "forward" : forward_list
            "email_closed" : email_closed_list
            }
        """

        emails_details = self.read_mail_details()  # Gets the subjects, urls, senders, dates, and labels from the main page
        emails_subjects = emails_details["subject"][0]
        emails_urls = emails_details["subject"][1]
        mails_senders = emails_details["sender"]
        mails_dates = emails_details["date"]
        mail_labels = emails_details["label"]

        # Create empty list that will be filled with the info that will be writen in the report: current_emails,
        # labels_list, subject_list, status_list, labels_check_text_list, url_list
        current_emails = []
        labels_list = []
        subject_list = []
        status_list = []
        labels_check_text_list = []
        url_list = []
        receiver_list = []
        client_list = []
        clients_check_list = []
        forward_list = []
        email_closed_list = []

        # Iterate each e-mail and get the details for each of them
        for subject_text, subject_url, sender, date, label in zip(emails_subjects,
                                                                  emails_urls,
                                                                  mails_senders, mails_dates,
                                                                  mail_labels):

            # If current e-mail date is the same as the last subject and date from the report, stop
            if self.last_mail_date:
                if date <= self.last_mail_date:
                    self.log.info("Stoped because " + str(date) + " is lower or equal than last report date " +
                                  str(self.last_mail_date))
                    print("Stoped because " + str(date) + " is lower or equal than last report date " +
                                  str(self.last_mail_date))
                    self.stop = True
                    break

            # If the mail has the label already set, skip to next mail. todo comment this for testing
            if label != '':
                labels_list.append(label)
                subject_list.append(subject_text)
                current_emails.append('')
                status_list.append("Labels were set before")  # todo check if page was loaded twice
                labels_check_text_list.append("")
                url_list.append("")
                receiver_list.append("")
                client_list.append("")
                clients_check_list.append("")
                forward_list.append("")
                email_closed_list.append("")
                continue

            # Get the link for the email body, get the rule by checking the body and set the label
            email_body_url = subject_url.replace("#mail", r"/show?msg")
            email_forward_url = subject_url.replace("#mail=", r"/compose/direction/forward/messageId/")
            # email_body_url = subject_url

            # https://leaseplnlu.icontroller.eu/messages/show?msg=155171
            # https://leaseplnlu.icontroller.eu/messages#mail=155171
            # https://leaseplangroup.icontroller.eu/messages/compose/direction/forward/messageId/721541

            # If no label is set for the current email it enters email body and returns its label(or None),
            # body, check text and subject
            email_info_dict = self.open_body(email_link_element=email_body_url, forward_url=email_forward_url)
            label, current_email, labels_check_text, subject_text, receiver_text, client, clients_check_text,\
            forward_address, email_closed = \
                email_info_dict["label"], email_info_dict["current_email"], email_info_dict["labels_check_text"], \
                email_info_dict["subject_text"], email_info_dict["receiver_text"], email_info_dict["client"], \
                email_info_dict["clients_check_text"], email_info_dict["forward"], email_info_dict["email_closed"]

            # Add label to labels list and add status
            if label:
                labels_list.append(label)
                labels_check_text_list.append(labels_check_text)
                status = "Label was added"
            else:
                labels_list.append("")
                labels_check_text_list.append("")
                status = "Couldn't set the label"

            subject_list.append(subject_text)

            self.log.info("The subject, sender, date, label for current mail are:-@" +
                          subject_text.encode("utf-8").decode() + "-@" + sender + "-@" + date + "-@" + str(label) +
                          "-@" + "-@")

            # Add info to lists
            current_emails.append(current_email)
            status_list.append(status)
            url_list.append(email_body_url)
            client_list.append(client)
            clients_check_list.append(clients_check_text)
            forward_list.append(forward_address)
            email_closed_list.append(email_closed)

        # Add current emails and labels to email details dictionary
        # emails_details["subject"] = emails_details["subject"][0]
        emails_details["subject"] = subject_list
        emails_details["label"] = labels_list
        emails_details["status"] = status_list
        emails_details["email_body"] = current_emails
        emails_details["check"] = labels_check_text_list
        emails_details["url"] = url_list
        emails_details["client"] = client_list
        emails_details["clients_check_text"] = clients_check_list
        emails_details["forward"] = forward_list
        emails_details["email_closed"] = email_closed_list



        # If we have last_mail_date, we should slice all the elements from the dictionary to that
        # mail
        if self.last_mail_date:
            where_to_cut = len(emails_details["label"])  # Labels list contains labels till last_mail_date
            for detail in emails_details:
                emails_details[detail] = emails_details[detail][:where_to_cut]

        if not (len(emails_details["subject"]) == len(emails_details["label"]) == len(emails_details["status"]) == len(
                emails_details["email_body"]) ==  len(emails_details["check"]) == len(emails_details["url"]) == len(
                emails_details["client"]) == len(emails_details["clients_check_text"]) == len(
                emails_details["forward"]) == len(emails_details["email_closed"])):
            for key in emails_details:
                print(emails_details[key])

        return emails_details

    def open_body(self, email_link_element, forward_url):
        """
        Opens the mail in a new and switches the focus to it. Based on 'subject', 'body' and 'sender' it gets the
        label and client if any rule matched. If label and client were found, it writes the values in the fields.
        Also forwards the email if forward address is available for the mached rule.

        Parameters
        ----------
        email_link_element : str
            The mail url passed from main page. It is used for opening the mail
        forward_url : The url of forward pop-up

        Returns
        -------
        A dictionary {"label": label, "current_email": current_email, "labels_check_text": labels_check_text,
                "subject_text": subject_text, "receiver_text": receiver_text, "client": client,
                "clients_check_text": clients_check_text, "forward": forward_address, "email_closed": email_closed}
        """

        # Find parent handle -> Main Window
        parent_handle = self.driver.current_window_handle

        # Find open window button and click it
        self.driver.execute_script('''window.open("{}","_blank");'''.format(email_link_element))
        # time.sleep(2)

        # Find all handles, there should two handles after clicking open window button
        handles = self.driver.window_handles

        label, current_email, labels_check_text, subject_text, receiver_text, client, clients_check_text, \
        forward_address, email_closed = 9 * " "
        # Switch to email body window
        for handle in handles:
            if handle not in parent_handle:
                self.driver.switch_to.window(handle)
                email_body_handle = self.driver.current_window_handle
                try:
                    current_email = self.get_current_email()
                except Exception as e:
                    current_email = ''
                    self.log.debug("Couldn't get the email body -- " + str(e))
                    screen_shot(self.driver, self.log)
                    send_error_email(self.error_receiver)
                    print(str(e))

                subject_text = self.get_text(locator=self._subject_from_popup_locator,
                                                     locator_type="css",
                                                     refreshes=1)

                sender_text = self.get_text(locator=self._sender_from_popup_locator,
                                                    locator_type="css")
                sender_text = sender_text[sender_text.find("<") + 1:].rstrip(">")

                receiver_text = self.get_text(locator=self._receiver_from_popup_locator,
                                                      locator_type="css")
                receiver_text = receiver_text[receiver_text.find("<") + 1:].rstrip(">")

                email_items = {'subject': subject_text, 'body': current_email,
                               'sender': sender_text, "receiver": receiver_text}

                label_email_info_dict = self.label_rules_obj.get_labels_dict(email_details=email_items)
                labels_check_text = label_email_info_dict["check_text"]
                # for test -> label_email_info_dict =
                # {'check_text': "Label in text", 'labels': "test label", 'forward': "yes", 'close': "yes"}
                start_get_client = time.time()
                client_email_info_dict = self.clients_rules_obj.get_labels_dict(email_details=email_items)
                clients_check_text = client_email_info_dict["check_text"]
                stop_get_client = time.time()
                print(f"It took {stop_get_client - start_get_client} seconds to get client")
                # for test -> client_email_info_dict =
                # {'check_text': "Client is ok", 'labels': "1111", 'forward': "yes", 'close': "yes"}
                #client_email_info_dict["labels"] = "1243"  # todo delete after testing

                label = self.set_label(to_fill=label_email_info_dict["labels"], is_client=False)

                if label:
                    #label_email_info_dict["labels"] = "test"  # todo delete after testing
                    client = self.set_label(to_fill=client_email_info_dict["labels"], is_client=True)

                # forward email todo create separate method for forward and close and test it
                #label_email_info_dict["forward"] = "victor.stanescu@leaseplan.com"  # todo delete after testing
                forward_address = label_email_info_dict["forward"]
                if client and label and "@" in str(forward_address):
                    self.driver.execute_script('''window.open("{}","_blank");'''.format(forward_url))
                    handles = self.driver.window_handles
                    for new_handle in handles:
                        if new_handle not in [parent_handle, email_body_handle]:
                            self.driver.switch_to.window(new_handle)
                            self.send_Keys(forward_address, self._forward_field_locator, "css")
                            self.send_Keys(Keys.ESCAPE, self._forward_field_locator, "css")
                            if not self.is_element_present(self._forward_accepted_locator, "css"):
                                forward_address = "Failed to forward"
                                self.driver.close()
                            else:
                                self.element_click(self._forward_button_locator, "css")
                                self.log.info("The email was forwarded")
                                print("Email was forwarded")
                else:
                    forward_address = "Not forwarded"
                self.driver.switch_to.window(email_body_handle)
                #label_email_info_dict["close"] = "yes"  # todo delete after testing
                if str(label_email_info_dict["close"]).strip().lower() == "yes" and client and \
                        label:
                    self.element_click(self._done_button_locator, "css") #todo uncomment for prod
                    email_closed = "Yes"
                    print("Email closed")
                else:
                    email_closed = "No"
                    self.driver.close()
                break

        # Switch back to the parent handle
        self.driver.switch_to.window(parent_handle)
        return {"label": label, "current_email": current_email, "labels_check_text": labels_check_text,
                "subject_text": subject_text, "receiver_text": receiver_text, "client": client,
                "clients_check_text": clients_check_text, "forward": forward_address, "email_closed": email_closed}

    def set_label(self, to_fill, is_client):
        """
        Fills the text in label or client field.
         Parameters
        ----------
        to_fill: Text to be filled
        is_client: If true, the text will be field in clients field, if false, label filed will be used

        """
        if is_client:
            field_locator = self._clients_field_locator
            locator_type = "css"
        else:
            field_locator = self._label_field_locator
            locator_type = "id"

        # check if labels field is empty
        label_present = self.is_element_present(self._label_body_locator, "css", False)
        if is_client:
            label_present = self.is_element_present("#s2id_autogen1 > ul > li.select2-search-choice > div", "css", False)

        if to_fill and not label_present:
            to_fill = to_fill.strip()
            # Picks a label if the rule contains more than one label and remove blank spaces
            self.send_Keys(to_fill, field_locator, locator_type) # todo uncomment for testing

            if is_client: # check if client number is in first dropdown list element
                time.sleep(1)
                client_dropdown_label = self.wait_for_element(locator=self._client_dropdown_label_locator,
                                                              locator_type="css").text
                if to_fill in client_dropdown_label:
                    time.sleep(0.5)
                    self.send_Keys(Keys.RETURN, field_locator, locator_type) # todo uncomment for testing
                    print("Client filled")
                else:
                    self.send_Keys(Keys.ESCAPE, field_locator, locator_type)
                    print("Failed to fill client")
                    to_fill = None
            else:
                self.send_Keys(Keys.RETURN, field_locator, locator_type) # todo uncomment for testing

                print("Label was set " + to_fill)
            time.sleep(0.5)

            self.send_Keys(Keys.ESCAPE, field_locator, locator_type)  # todo comment for testing
            # Check if label was acceped (is in labels list)
            label_accepted = self.is_element_present(locator="div.labels li.select2-search-choice",
                                                    locatorType="css",
                                                    take_screen_shot=False) # todo comment for testing
            if is_client:
                label_accepted = self.is_element_present(locator="#s2id_autogen1 > ul > li.select2-search-choice",
                                                         locatorType="css",
                                                         take_screen_shot=False)
            #self.send_Keys(Keys.BACK_SPACE, field_locator, locator_type) # todo uncomment for testing
            #self.send_Keys(Keys.BACK_SPACE, field_locator, locator_type) # todo uncomment for testing
            #self.send_Keys(Keys.ESCAPE, field_locator, locator_type)  # todo uncomment for testing
            #label_accepted = True # todo uncomment for testing
            if not label_accepted:
                to_fill = None
        else:
            to_fill = None
        return to_fill

    def read_mail_details(self):
        """
        If page has loaded, it returns a dictionary that contains the e-mails details from main page

        Returns
        -------
        {
        "subject" : (mails_subject_text_list, mails_subject_urls)
        "sender" : mails_sender_text_list
        "date" : mails_date_text_list
        "label" : mails_label_text_list
        }
        """

        mail_details = {}

        mails_subject_elements = self.wait_for_element(locator=self._subject_locator,
                                                       locator_type="css",
                                                       condition_text="presence_of_all_elements_located",
                                                       check_page_load=True,
                                                       refreshes=3)
        mails_sender_elements = self.wait_for_element(locator=self._sender_locator,
                                                      locator_type="css",
                                                      condition_text="presence_of_all_elements_located",
                                                      check_page_load=True)
        mails_date_elements = self.wait_for_element(locator=self._date_locator,
                                                      locator_type="css",
                                                      condition_text="presence_of_all_elements_located",
                                                      check_page_load=True)
        mails_labels_elements = self.wait_for_element(locator=self._label_locator,
                                                      locator_type="css",
                                                      condition_text="presence_of_all_elements_located",
                                                      check_page_load=True)

        # Convert mails objects to text
        # todo refactor get lists Use table.py try to get it in pandas
        mails_subject_text_list = [mail_subject.text for mail_subject in mails_subject_elements]
        mails_subject_urls = [mail_subject.get_attribute("href") for mail_subject in mails_subject_elements]
        mails_sender_text_list = [mail_sender.text for mail_sender in mails_sender_elements]
        mails_date_text_list = [mail_date.text for mail_date in mails_date_elements]
        mails_label_text_list = [mail_label.text for mail_label in mails_labels_elements]

        # Add detailes to the dictionary
        mail_details["subject"] = (mails_subject_text_list, mails_subject_urls)
        mail_details["sender"] = mails_sender_text_list
        mail_details["date"] = mails_date_text_list
        mail_details["label"] = mails_label_text_list

        return mail_details
        # mails_details = {"subjects":(mails_subject_text_list, mails_subject_elements),
        #                   "sender": mails_sender_text_list,
        #                   "date": mails_date_text_list}

    def get_current_email(self):
        """
        Extracts the current email, eliminating the previous replies from the email body. It searches for
        predefined word like ("From: ", ) and gets the text from the beginning to the found word. If no word was found,
        it returns the whole body

        Returns
        -------
        Current mail string
        """

        # The email that contains full conversation (initial, replys, etc)
        email_body = self.wait_for_element(locator=self._email_body_locator,
                                           locator_type="css", timeout=20, refreshes=1).text

        for word in self.sender_words: # todo check if two words are in the same mail
            if email_body.lower().find(word.lower()) != -1:
                current_email = email_body[:email_body.find(word)]
                for next_word in self.next_email_words:
                    if next_word in current_email:
                        next_emails_text = email_body[email_body.find(word)+len(word):]
                        current_email = next_emails_text[:next_emails_text.find(word)]
                        break
                break
            # if email_body.find(word) != -1:
            #     current_email = email_body[:email_body.find(word)]
            #     break
        else:
            current_email = email_body[:email_body.find(word)]
        return current_email



    def close_browser(self):
        self.driver.close()



if __name__ == "__main__":
    from app.base.webdriverfactory import WebDriverFactory
    from app.pages.login_page import LoginPage

    driver = WebDriverFactory("chrome", "https://leaseplangroup.icontroller.eu/messages").getWebDriverInstance()
    LoginPage(driver).login()
    labels_file = r"D:\Users\eddiehamilton\LeasePlan Information Services\LPSC - Eagle - RPA - IController\Rotulos.xlsx"
    clients_file = r"D:\Users\eddiehamilton\LeasePlan Information Services\LPSC - Eagle - RPA - IController\Clientes.xlsx"
    iControler_auto_obj = WebInterations(driver, label_rules_obj=labels_file, clients_rules_obj=clients_file)
    iControler_auto_obj.get_emails_details()
    #iControler_auto_obj.run_automation_test()