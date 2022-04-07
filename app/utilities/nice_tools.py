# Created by Victor Stanescu at 2/6/2019

# Enter feature description here

# Enter steps here
import sys
import traceback

import win32com.client as win32
import os

class MailUtils:
    """Class for reading, writing emails"""

    def __init__(self):
        self.outlook = win32.Dispatch('outlook.application')

    def create_mail(self, to, subject, cc=None, body='', html_body=None, attachments=None):
        """
        Creates a mail object that has the main elements: sender, receiver, cc, subject, body, attachments

        Parameters
        ----------
        to : str
            Receiver
        subject: str
            Mail subject
        cc : str
            CC addresses
        body : str
            Mail body
        html_body : str
            Mail body in html format
        attachments : str, list
            Mail attachments

        Returns
        -------
            Mail object

        """
        mail = self.outlook.CreateItem(0)
        mail.To = to
        if cc:
            mail.CC = cc
        mail.Subject = subject
        if html_body:
            mail.HTMLBody = html_body
        if body != '': #todo remove this
            mail.Body = body
        if type(attachments) == str:
            if os.path.isdir(attachments):
                attachments = [os.path.join(attachments, att) for att in os.listdir(attachments)]
            elif os.path.isfile(attachments):
                mail.Attachments.Add(os.path.abspath(attachments))
            else:
                raise TypeError("You should put the absolute path of the file or folder in attachments")

        if type(attachments) == list:
            for attachment in attachments:
                mail.Attachments.Add(os.path.abspath(attachment))
        return mail

    def send_mail(self, mail):
        """
        Sends an email using a mail object made with create_mail metrhod

        Parameters
        ----------
        mail : mail object
        A mail object created by create_mail method
        """
        mail.Send()


class GetPaths:
    """
    For testing. Gets the paths used in automation
    """
    @staticmethod
    def get_report_folder(country):
        reports_folder_path = os.path.join((os.path.split(sys.argv[0])[0]), "reports_" + country)
        return reports_folder_path

    @staticmethod
    def get_report_file_path(report_folder_path, date, country):
        date_today = str(date.today())
        report_path = os.path.join((os.path.split(sys.argv[0])[0]), report_folder_path,
                                   "report_{}_{}.xlsx".format(country, date_today))
        return report_path

    @staticmethod
    def get_config_file(): #todo move from here
        """
        Config file should be in configFiles. This file should be send with the daily report

        Returns
        -------
        Config file path
        """
        #config_folder = os.path.abspath(sys.argv[0] + "/../../configFiles")
        config_folder = os.path.curdir

        labels_cfg_file_path = None

        for file in os.listdir(config_folder):
            if "Labels_cfg" in file:
                labels_cfg_file_path = os.path.join(config_folder, file)
                break
        else:
            assert (labels_cfg_file_path is not None), "Couldn't find the labels config file"
        return labels_cfg_file_path

    @staticmethod
    def pie_path(report_file):
        os.path.join(os.path.dirname(report_file), "Pies")
        return

    @staticmethod
    def newest_file(folder_path):
        """
        Gets the last file created/modified in a filder

        Parameters
        ----------
        folder_path : str
            Path of the folder

        Returns
        -------
            The last file that has been created/modified in the input folder
        """
        files = os.listdir(folder_path)
        files_cleaned = list(filter(
            lambda file: "backup" not in file and os.path.isfile(os.path.join(folder_path, file)), files)
        )
        paths = [os.path.join(folder_path, basename) for basename in files_cleaned]
        if len(paths) > 0:
            newest_file_path = max(paths, key=os.path.getctime)
        else:
            newest_file_path = None
        return newest_file_path


def newest_file(folder_path):
    """
    Gets the last file created/modified in a filder

    Parameters
    ----------
    folder_path : str
        Path of the folder

    Returns
    -------
        The last file that has been created/modified in the input folder
    """
    files = os.listdir(folder_path)
    files_cleaned = list(filter(
        lambda file: "backup" not in file and os.path.isfile(os.path.join(folder_path, file)), files)
    )
    paths = [os.path.join(folder_path, basename) for basename in files_cleaned]
    if len(paths) > 0:
        newest_file_path = max(paths, key=os.path.getctime)
    else:
        newest_file_path = None
    return newest_file_path

def send_error_email(error_receiver_address, excepttion_message=''):
    error_msg = 'This is the exception:\n {} \n\n Exception message: \n{}'.format(traceback.format_exc(),
                                                                                  excepttion_message)
    print(traceback.format_exc())

    mail_util = MailUtils()
    mail = mail_util.create_mail(to=error_receiver_address,
                                 subject="[iController PT] Couldn't run app",
                                 body=error_msg)
    mail_util.send_mail(mail)

if __name__ == "__main__":
    mail_util = MailUtils()
    mail = mail_util.create_mail("victor.stanescu@leaseplan.com",
                                 "sss",
                                 "bbb",
                                 r"C:\Users\stanv\AppData\Local\Temp\gen_py\3.7")
    mail_util.send_mail(mail)

