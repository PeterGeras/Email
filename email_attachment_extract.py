# Import Libs
import sys
import imaplib
import getpass
import email
import email.header
from datetime import datetime
import re
import base64
import uuid
import os
import shutil
import cx_Oracle

from email.message import EmailMessage
from email.header import decode_header

# Python file
import file_move_helper
import config

import logzero
import logging
from logzero import logger


# Setup rotating logfile with 3 rotations, each with a maximum filesize of 1MB:
log_file = os.path.join('logs', 'inbound_email.log')
logzero.logfile(log_file, maxBytes=1e6, backupCount=3)
# Set a minimum log level
logzero.loglevel(logging.INFO)
# Set a custom formatter
formatter = logging.Formatter('%(asctime)-15s - %(levelname)s: %(message)s')
logzero.formatter(formatter)

# Email Config
EMAIL_ACCOUNT           = config.email_prod['account']
EMAIL_PASSWORD          = config.email_prod['password']
EMAIL_FOLDER            = config.email_prod['folder']
EMAIL_COMPLETE_FOLDER   = config.email_prod['folder_complete']
EMAIL_FAIL_FOLDER       = config.email_prod['folder_fail']

# Locations
cwd = os.getcwd()
STAGING_DIRECTORY = file_move_helper.INPUT_DIRECTORY
OUTPUT_DIRECTORY_NORMAL = file_move_helper.OUTPUT_DIRECTORY_NORMAL
OUTPUT_DIRECTORY_DEFENCE = file_move_helper.OUTPUT_DIRECTORY_DEFENCE
OUTPUT_DIRECTORY_FAILURE = file_move_helper.OUTPUT_DIRECTORY_FAILURE

# Oracle Config
username        = config.oracle['username']
password        = config.oracle['password']
databaseIP      = config.oracle['databaseIP']
oracleSchema    = config.oracle['schema']

# Connect to Oracle
connection = cx_Oracle.connect(username, password, databaseIP)
connection.current_schema = oracleSchema
cursor = connection.cursor()


def CleanCharacters(fileName):

    if fileName[:11] == '=?KOI8-R?B?':
        fileName = base64.b64decode(fileName[11:]).decode('KOI8-R') 
    if fileName[:10] == '=?UTF-8?B?':
        fileName = base64.b64decode(fileName[10:]).decode('utf-8')
    if fileName[:10] == '=?UTF-8?Q?':
        fileName = quopri.decodestring(fileName[10:]).decode('utf-8')
        fileName=fileName.replace("?","")
        fileName=fileName.replace("PO","")


    """
    Apply Regular Expression to remove reserved characters for Windows Filenames
    """
    # Special Charaters that need to removed from the email subject
    # < (less than)
    # > (greater than)
    # : (colon)
    # " (double quote)
    # / (forward slash)
    # \ (backslash)
    # | (vertical bar or pipe)
    # ? (question mark)
    # * (asterisk)
    decoded = fileName.encode("ascii", "ignore").decode()
    clean = re.sub('[<>:"/\\|?*]', '', decoded)
    if fileName != clean:
        logger.info('Filename updated from %s to %s' % (fileName, clean))

    return clean


# Check Manhattan based on PO number
def checkPOManhattan(PO):
    cursor_named_params = {'PO_ORDER': PO}

    cursor.execute("""
    SELECT PO_ORDER
    FROM ORDERH
    WHERE PO_ORDER = CAST(TRIM(:PO_ORDER) as NCHAR(20))
    """, cursor_named_params)

    results = cursor.fetchone()

    if results:
        logger.info("PO exists in Manhattan")
        return True
    else:
        logger.info("PO does not exist in Manhattan")
        return False


def msg_walk(msg):
    validEmail = False

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            # print part.as_string()
            continue
        if part.get('Content-Disposition') is None:
            # print part.as_string()
            continue
        fileName = part.get_filename()
        # print ("  ->Attachment FileName:")
        # print (fileName)

        if bool(fileName):
            if fileName[:11] == '=?KOI8-R?B?':
                fileName = base64.b64decode(fileName[11:]).decode('KOI8-R')
            if fileName[:10] == '=?UTF-8?B?':
                # print ("UTF->")
                fileName = base64.b64decode(fileName[10:])
                b = fileName
                fileName = b.decode('utf-8')

            filename = CleanCharacters(str(fileName))
            logger.info("\t->Attachment FileName:" + filename)

            # Use Regular expression to break up PO and the Document type
            # SDKT – Service Docket
            # RCA – Root Cause Analysis
            # SWMS – Safe Work Method Statement
            # ECBD – Estimated Cost Breakdown
            attachmentPattern = r"(?i)(\d+)-(sdkt|rca|swms|ecbd|comp).(\w+)"
            attachmentMatch = re.search(attachmentPattern, filename)
            if attachmentMatch:
                logger.info("\tAttachment pattern satisfied")
                PO = attachmentMatch.group(1)
                documentType = attachmentMatch.group(2)

                # Login to Oracle and check Manhattan
                PO_valid = checkPOManhattan(PO)

                if PO_valid:
                    validEmail = True
                    # stage the file for upload to SharePoint
                    filePath = os.path.join(STAGING_DIRECTORY, filename)
                    if not os.path.isfile(filePath):
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                else:
                    validEmail = False
                    # break  # Break out of attachment loop when invalid PO
            else:
                logger.info("\tAttachment pattern failed")

        else:
            validEmail = False
            # break  # Break out of attachment loop when does not match regular expression

    return validEmail


def process_PO(MailItem):
    """
    Do something with emails messages in the folder.
    For the sake of this example, print some headers.
    """

    rv, data = MailItem.search(None, "ALL")
    if rv != 'OK':
        logger.info("No messages found!")
        return
    for num in data[0].split():
        rv, data = MailItem.fetch(num, '(RFC822)')
        if rv != 'OK':
            logger.info("ERROR getting message", num)
            return
        msg = email.message_from_bytes(data[0][1])
        decode = email.header.decode_header(msg['Subject'])[0]
        subject = str(decode[0])
        decode = email.header.decode_header(msg['From'])[0]
        EmailFrom = str(decode[0])

        logger.info("Email From:" + EmailFrom)
        # print ("Subject   :" + subject)
        # print (email.header)
        # print (get_part_filename(msg))

        # download each file
        # if no file - move to exception
        # for each file Regular expression and check for PO number validity flag as false as soon as email has invalid content

        # loop through each attachment in the email and save valid attachments to STAGING_DIRECTORY
        validEmail = msg_walk(msg)

        if validEmail:
            rv, resp = MailItem.copy(num,EMAIL_COMPLETE_FOLDER)
            #Revisit
            if rv == 'OK':
                MailItem.store(num, '+FLAGS', r'(\Deleted)') 
        else:
            rv, resp = MailItem.copy(num,EMAIL_FAIL_FOLDER)
            #Revisit
            if rv == 'OK':
                MailItem.store(num, '+FLAGS', r'(\Deleted)')


def main():
    startTime = datetime.now()

    M = imaplib.IMAP4_SSL('outlook.office365.com')

    try:
        rv, data = M.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
    except imaplib.IMAP4.error:
        logger.info("LOGIN FAILED!!!")
        return False

    logger.info(rv)
    logger.info(data)
    # print (rv, data)

    rv, mailboxes = M.list()
    if rv == 'OK':
        logger.info("Mailboxes:")
    # print (mailboxes)

    rv, data = M.select(EMAIL_FOLDER)
    if rv == 'OK':
        logger.info("Processing mailbox...\n")
        # process_mailbox(M)
        process_PO(M)
        M.close()
    else:
        logger.info("ERROR: Unable to open mailbox: " + rv)
        
    # Move staging directory files to required folders
    file_move_helper.main()

    M.logout()
    
    logger.info("Program runtime: " + str(datetime.now() - startTime) + "s\n")

    return True


if __name__ == '__main__': main()

