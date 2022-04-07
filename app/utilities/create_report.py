import pandas as pd
import os
from shutil import copyfile
import matplotlib.pyplot as plt
from datetime import date
import re
from app.utilities.database import CursorFromConnectionFromPool
from app.utilities.database import Database
import logging
from app.utilities import custom_logger as cl
from app.utilities import nice_tools


class CreateReport(object):
    """
    Class used for creating and manipulating reports

    Parameters
    ----------
    _dates_column : str
        Column name used to get last date
    _subject_column : str
        Column name used to get last subject
    file_path : str
        Path for the current (today) report. If it doesn't exists, it will be created
    last_report_path : str
        Path to previous report
    sheet_name : str
        Name of the reports sheet

    Methods
    -------
    backup_report()
        Makes a backup for the current report
    get_last_email_date_and_subj()
        Returns the date and the subject from the previous report
    append_df_to_excel(data)
        Appends a dataframe to excel file and saves it. Used in write_data_to_excel method
    write_data_to_excel(data=None)
        Appends data to an existing report. If it doesn't exists, it creates one
    raport_to_html(df)
        Converts data to  a html table
    filter_report()
        Filters a dataframe by a given condition.
    get_info_from_filename()
        Return the text between "_" and "." from a report filename
    create_pie()
        Creates a pie graph for a given field
    """

    log = cl.customLogger(logging.INFO)
    _dates_column = "date"
    _subject_column = "subject"

    def __init__(self, file_path, last_report_path=None, sheet_name="Sheet1"):
        """
        Creates a report instance

        Parameters
        ----------
        file_path : str
            Path for the current (today) report. If it doesn't exists, it will be created
        last_report_path : str
            Path to previous report
        sheet_name : str
            Name of the reports sheet
        """
        self.file_path = file_path
        self.pies_folder = os.path.join(os.path.dirname(self.file_path), "Pies")

        self.last_report_path = last_report_path
        self.backup_excel_file = self.file_path[:self.file_path.rfind(".")] + "_backup.xlsx"
        self.sheet_name = sheet_name

        self.df_today = None
        self.df_last_report = pd.DataFrame()

        if os.path.exists(self.file_path):
            self.df_today = pd.read_excel(self.file_path)  # todo sort
            self.df_today.sort_values("date", ascending=False, inplace=True)
            self.df_today.reset_index(drop=True, inplace=True)
        elif self.last_report_path is not None:
            print("Today report dosen't exist. It will be created: " + self.file_path)
            if os.path.exists(self.last_report_path):
                self.df_last_report = pd.read_excel(self.last_report_path)
                self.df_last_report.sort_values("date", ascending=False, inplace=True)
                self.df_last_report.reset_index(drop=True, inplace=True)
                # todo send email when here
                print("Raport for today doesn't exists")

    def backup_report(self):
        """
        Makes a backup for the current report

        Returns
        -------
        None
        """
        if os.path.exists(self.backup_excel_file):
            os.remove(self.backup_excel_file)
            copyfile(self.file_path, self.backup_excel_file)
        else:
            copyfile(self.file_path, self.backup_excel_file)

    def get_last_email_date_and_subj(self):
        """
        Returns the date and the subject from the previous report

        Returns
        -------
        The date and the subject from the previous report
        """
        if self.df_today is not None and not self.df_today.empty:
            last_mail_date = str(self.df_today.at[0, self._dates_column])
            last_mail_subject = self.df_today.at[0, self._subject_column]
            return [last_mail_subject, last_mail_date]
        elif self.df_last_report is not None and not self.df_last_report.empty:
            last_mail_date = str(self.df_last_report.at[0, self._dates_column])
            last_mail_subject = self.df_last_report.at[0, self._subject_column]
            print("Couldn't find the report for today. the report from last day will"
                  " be used for start processing time: " + last_mail_date)
            return [last_mail_subject, last_mail_date]
        else:
            # todo send mail here
            print("No report was found in reports folder")
            return [None, None]

    def append_df_to_excel(self, data):
        """
        Appends a dataframe to excel file and saves it. Used in write_data_to_excel method

        Parameters
        ----------
        data : pandas dataframe
            Dataframe to be appended

        Returns
        -------
        None
        """
        result = pd.concat([self.df_today, data], ignore_index=True)
        result.drop_duplicates(inplace=True)
        result.to_excel(self.file_path, index=False)

    def write_data_to_excel(self, data=None):
        """
        Appends data to an existing report. If it doesn't exists, it creates one

        Parameters
        ----------
        data : pandas dataframe
            Dataframe to be appended

        Returns
        -------
        None
        """
        if os.path.exists(self.file_path):
            self.append_df_to_excel(data)
        else:
            data.to_excel(self.file_path, index=False)

    def write_data_to_db(self, data, db_table, db_table_columns):
        try:
            assert (data.shape[1] == len(db_table_columns.split(","))), "Number of input columns is not equal to db " \
                                                                        "table columns"
            for row in range(data.shape[0]):
                values = ", ".join(data.iloc[row].values.astype(str))
                insert_statement = f"INSERT INTO {db_table} ({db_table_columns}) VALUES ({values})"
                with CursorFromConnectionFromPool() as cursor:
                    cursor.execute(insert_statement)
        except Exception as e:
            self.log.error("Failed to add to db: " + repr(e))

    def excel_to_db(self, db_table, db_table_columns, error_receiver_address='victor.stanescu@leaseplan.com'):
        try:
            dt = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
        except Exception as e:
            print("Failed to read excel: " + repr(e))
            raise
        try:
            Database.initialise()
            excel_columns = dt.columns.values
            assert (len(excel_columns == len(db_table_columns.split(",")))), "Number of input columns is not equal " \
                                                                             "to db table columns"
            excel_columns_text = ", ".join(excel_columns)
            # print("excel columns: " + excel_columns_text)
            # print("Db columns: " + db_table_columns)
            for row in range(dt.shape[0]):
                clean_row = [i.replace("'", "''") for i in dt.iloc[row].values.astype(str)]
                values = [f"'{i}'" for i in clean_row]
                #print("No of values: " + str(len(values)))
                #print("No of columns: " + str(len(db_table_columns.split(","))))
                values = ", ".join(values)
                insert_statement = f'INSERT INTO "{db_table}" ({db_table_columns}) VALUES ({values})'
                with CursorFromConnectionFromPool() as cursor:
                    cursor.execute(insert_statement)
        except Exception as e:
            self.log.error("Failed to add to db: " + repr(e))
            print("Failed to add to db: " + repr(e))
            print("SQL: " + insert_statement)
            nice_tools.send_error_email(error_receiver_address)


    @staticmethod
    def raport_to_html(df):
        """
        Converts data to  a html table
        Parameters
        ----------
        df : pandas dataframe
            Dataframe to be converted
        Returns
        -------
            Converted html table
        """
        return df.to_html()

    def filter_report(self, filter_condition, add_to_index=None):
        """
        Filters a dataframe by a given condition.
        Parameters
        ----------
        filter_condition : str
            Condition for filtering. Check https://pbpython.com/excel-filter-edit.html for info about creating string
        add_to_index
            If dataframe has a header, a number of rows can be skiped by this number
        Returns
        -------
            Filtered dataframe
        """
        filtered_df = eval("self.df_today.loc[self.df_today" + filter_condition + "]")
        if add_to_index:
            filtered_df.index = filtered_df.index + add_to_index
        return filtered_df

    @staticmethod
    def get_info_from_filename(file):
        """
        Return the text between "_" and "." from a report filename
        Parameters
        ----------
        file : str
            File path
        Returns
        -------
            The text between "_" and "." from a report filename
        """
        base_name = os.path.basename(file)
        return base_name[base_name.find("_") + 1:base_name.find(".")]

    def create_pie(self, report_field="status", title="Labels status"):
        """
        Creates a pie graph for a given field
        Parameters
        ----------
        report_field : str
            The field used for creating the pie
        title : str
            Pie title
        Returns
        -------
            Pie path
        """
        plt.style.use("ggplot")
        plt.clf()
        self.df_today[report_field].value_counts().plot(kind='pie', startangle=0, title=title, autopct='%1.1f%%')
        plt.axis("off")
        # plt.show()
        if not os.path.exists(self.pies_folder):
            os.mkdir(self.pies_folder)
        date_and_country = self.get_info_from_filename(self.file_path)
        pie_name = "pie_" + date_and_country + ".png"
        pie_path = os.path.join(self.pies_folder, pie_name)
        plt.savefig(pie_path)
        return pie_path


class RecoverReport:
    """Recovers a report from logs using separators"""

    def __init__(self, log_file, separator, recovered_report_path):
        """
        Parameters
        ----------
        log_file : str
            file used to recover the report
        separator
            separator used to get report details
        recovered_report_path
            this report file will be created
        """
        self.log_file = log_file
        self.separator = separator
        self.labels = ["subject", "sender", "date", "label", "country"]
        # self.labels = ["subject", "sender", "date", "label"]
        self.recovered_report_path = recovered_report_path

    def create_report_from_logs(self):
        list_all = []
        with open(self.log_file) as rf:
            for line in rf:
                a = re.split(self.separator, line)
                d_list = tuple(a[1:-1])
                if len(d_list) > 0:
                    list_all.append(d_list)

        df = pd.DataFrame.from_records(list_all, columns=self.labels)
        df.drop_duplicates(inplace=True)
        df.to_excel(self.recovered_report_path, index=False)


class TablesUtil(object):

    def __init__(self):
        self.all_data = None  # The dataframe from previous page

    def data_from_dict(self, data_dict):
        new_df = pd.DataFrame.from_dict(data_dict)
        return new_df

    def append_data(self, current_df):
        result = pd.concat([self.all_data, current_df], ignore_index=True)
        result.drop_duplicates(inplace=True)
        return result


if __name__ == "__main__":

### Recover report
    # log_file = r"C:\projects\iController\app\main_app\iController_Logs\automation_2019-09-05.log"
    # recover_file = r"C:\Users\stanv\Desktop\recover.xlsx"
    # separator = "-@"
    # recover_obj = RecoverReport(log_file, separator, recover_file)
    # recover_obj.create_report_from_logs()

    report_file = r"D:\Victor\PT_iController\app\main_app\Reports\report_2020-10-22.xlsx"
#subject	sender	date	label	status	email_body	check	url	client	clients_check_text	forward	email_closed

    db_cols = "subject,sender,date,label,status,email_body,label_" \
                     "check,url,client,client_check,forward,email_closed"
    report = CreateReport(report_file)
    report.excel_to_db('"'"lpsc_pt_iController"'"', db_cols)