##########################################################
#
# Error emailing - Astea EAI Service Task
#
# Peter Geras
# Release History
# 03.04.2020	Initial Release
#
##########################################################


from datetime import datetime

import timestamp
from logger import logger


def emailing():
    import email_outlook
    sent = email_outlook.send_email()
    return sent


def program_runtime(start, end):
    diff_time = end - start
    minutes, seconds = divmod(diff_time.seconds, 60)
    microseconds = diff_time.microseconds
    milliseconds = microseconds//1000

    return seconds, milliseconds


def main():
    start_time = datetime.now()
    sent_email = False

    if timestamp.to_run_program():
        sent_email = emailing()

    end_time = datetime.now()

    if sent_email:
        seconds, milliseconds = program_runtime(start_time, end_time)
        timestamp.file_write(start_time)
        logger.info(f'Program run time: {seconds}.{milliseconds:03d}s\n')

    return


if __name__ == '__main__':
    main()
