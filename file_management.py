import os
import sys
import re
import shutil
from datetime import datetime, timedelta
from logger_updated_ipo_global_report import logger, log_folder, log_file


def delete_old_files(folder: str, num_days: int = 30) -> list:
    """
    Deletes files older than the number of days given as a parameter. Defaults to delete files more than 30 days old.
    :param folder: folder location files will be deleted from
    :param num_days: int specifying the number of days before a file is deleted
    :return: list of files that were deleted
    """
    old_date = datetime.utcnow() - timedelta(days=num_days)
    files_deleted = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            f_abs = os.path.join(root, file)
            f_modified = datetime.fromtimestamp(os.path.getmtime(f_abs))
            if f_modified <= old_date:
                os.unlink(f_abs)
                files_deleted.append(file)
    if len(files_deleted) > 0:
        logger.info(f"Deleted {', '.join(files_deleted)}")
    return files_deleted


def delete_old_files_test(folder: str, num_days: int = 30) -> list:
    """
    Used to test the delete_old_files function. This will only print the name of the files that would be deleted.
    :param folder: folder location files will be deleted from
    :param num_days: int specifying the number of days before a file is deleted
    :return: list of files that were deleted
    """
    old_date = datetime.utcnow() - timedelta(days=num_days)
    files_deleted = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            f_abs = os.path.join(root, file)
            f_modified = datetime.fromtimestamp(os.path.getmtime(f_abs))
            if f_modified <= old_date:
                print(f_abs)
                files_deleted.append(file)
    if len(files_deleted) > 0:
        print(f"Deleted {', '.join(files_deleted)}")
    return files_deleted


def archive_logs(num_days: int = 30):
    current_log_file = os.path.join(log_folder, log_file)
    with open(current_log_file, 'r') as f:
        all_lines = f.readlines()
    key_dates = {'first_log_date': return_date_str(all_lines[0]), 'last_log_date': return_date_str(all_lines[-1])}
    old_date = datetime.utcnow() - timedelta(days=num_days)
    if key_dates.get('first_log_date') is not None:
        first_log_date_datetime = datetime.strptime(key_dates.get('first_log_date'), '%Y-%m-%d')
        if first_log_date_datetime < old_date:
            log_file_name, log_file_ext = os.path.splitext(log_file)
            archived_log = f"{log_file_name} {key_dates.get('first_log_date')} - {key_dates.get('last_log_date', datetime.today().strftime('%Y-%m-%d'))}{log_file_ext}"
            shutil.move(current_log_file, os.path.join(log_folder, 'Previous Logs', archived_log))


def return_date_str(text: str, date_pat: str = r"(\d{4}\-\d{2}\-\d{2})"):
    mo = re.search(date_pat, text)
    if mo is not None:
        return mo.group()


def main():
    try:
        for folder in [os.path.join(os.getcwd(), 'Logs', 'Email Attachments'), os.path.join(os.getcwd(), 'Results')]:
            delete_old_files(folder)
        archive_logs()
    except Exception as e:
        logger.error(e, exc_info=sys.exc_info())


if __name__ == '__main__':
    main()
