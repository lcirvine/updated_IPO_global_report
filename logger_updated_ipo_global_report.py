import os
import logging
from datetime import date
import configparser
import win32com.client as win32


log_file = 'Updated IPO Global Report Logs.txt'
log_folder = os.path.join(os.getcwd(), 'Logs')
today_date = date.today().strftime('%Y-%m-%d')

if not os.path.exists(log_folder):
    os.mkdir(log_folder)
handler = logging.FileHandler(os.path.join(log_folder, log_file), mode='a+', encoding='UTF-8')
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger = logging.getLogger()
logger.addHandler(handler)
logger.setLevel(logging.INFO)


def error_email(error_message: str = ''):
    """
    Used to send an email when an error is encountered.
    Email details like sender and recipients are provided in .ini file which is read by configparser.
    :param error_message: optional string that will be added to body of email
    :return: None
    """
    config = configparser.ConfigParser()
    config.read('settings_update_ipo_report.ini')
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = config.get('email', 'errorTo')
    mail.Sender = config.get('email', 'sender')
    mail.Subject = f"ERROR: {config.get('email', 'subject')} {today_date}"
    mail.HTMLBody = config.get('email', 'errorBody') + error_message + config.get('email', 'signature')
    mail.Attachments.Add(os.path.join(log_folder, log_file))
    mail.Send()
