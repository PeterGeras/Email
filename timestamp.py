import os
from datetime import datetime

from logger import logger

# Folders and files
cwd = os.getcwd()
time_file = r'time_tracker.txt'

# Required minutes to wait until running program again
MINUTES_BETWEEN_EMAIL = 30


def file_read():
    try:
        with open(time_file, 'r') as f:
            date_text = f.read()
            date_object = datetime.strptime(date_text, '%Y-%m-%d %H:%M:%S.%f')
    except Exception as ex:
        logger.exception({ex})
        return False

    return date_object


def file_write(t):
    try:
        with open(time_file, 'w') as f:
            f.write(str(t))
    except Exception as ex:
        logger.exception({ex})
        return False

    return True


def to_run_program():
    start_time = datetime.now()

    logged_time = file_read()
    if logged_time is False:
        return False

    difference_time = start_time - logged_time
    minutes, seconds = divmod(difference_time.seconds, 60)

    msg_diff_minutes = f'Minutes since previous run: {minutes}'
    if minutes < MINUTES_BETWEEN_EMAIL:
        logger.info(f'{msg_diff_minutes} - requires {MINUTES_BETWEEN_EMAIL} minutes')
        return False

    return True
