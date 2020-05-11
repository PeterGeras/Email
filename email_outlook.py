import smtplib
import email.message

from logger import logger

# Email credentials
EMAIL_ACCOUNT = ''
EMAIL_PASSWORD = ''
smtp = smtplib.SMTP("smtp-mail.outlook.com", 587)

EMAILS_TO = [
    # '',
    ''
]


def content():
    msg = email.message.Message()
    msg['Subject'] = ""
    msg['From'] = EMAIL_ACCOUNT
    msg['To'] = ', '.join(EMAILS_TO)
    msg.add_header('Content-Type', 'text/html')

    body = """\
    <h1 style="color: #5e9ca0;"><span style="color: #800000;">ERROR</span></h1>
    <ol>
       <li>
          <h4><span style="color: #000000;">Event Viewer </span></h4>
       </li>
       <li>
          <h4><span style="color: #000000;">Applications and Services Logs</span></h4>
       </li>
    </ol>
    """

    msg.set_payload(body)

    return msg


def login():
    try:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    except Exception as ex:
        logger.error(f'Login failed - exception: {ex}')
        return False

    return True


def send_email():
    sent_success = False

    login_success = login()
    if login_success is True:
        msg = content()
        try:
            smtp.sendmail(msg['From'], EMAILS_TO, msg.as_string())  # Don't use msg['To'] as that's a string, not list
            logger.info(f'\t# EMAIL SENT #')
            sent_success = True
        except Exception as ex:
            logger.exception({ex})

    return sent_success
