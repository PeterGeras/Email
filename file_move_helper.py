# Import Libs
import sys
import os
import re
import timeit
import shutil
import cx_Oracle

# Python file
import config

import logzero
import logging
from logzero import logger

# Setup rotating logfile with 3 rotations, each with a maximum filesize of 1MB:
log_file = os.path.join('logs', 'file_move.log')
logzero.logfile(log_file, maxBytes=1e6, backupCount=3)
# Set a minimum log level
logzero.loglevel(logging.INFO)
# Set a custom formatter
formatter = logging.Formatter('%(asctime)-15s - %(levelname)s: %(message)s')
logzero.formatter(formatter)

# Locations
cwd = os.getcwd()
INPUT_DIRECTORY = os.path.join(cwd, 'staging')
OUTPUT_DIRECTORY_NORMAL = os.path.join(cwd, 'output_normal')
OUTPUT_DIRECTORY_DEFENCE = os.path.join(cwd, 'output_defence')
OUTPUT_DIRECTORY_FAILURE = os.path.join(cwd, 'output_failed')

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
        logger.info('Filename updated from %s to %s' % (dirty, clean))

    return clean


# Pick output folder
def output_directory(clientGroup):
    directory = OUTPUT_DIRECTORY_FAILURE

    try:
        client = str(clientGroup)
        
        if "DENNSW" in client:
            directory = OUTPUT_DIRECTORY_DEFENCE
        else:
            directory = OUTPUT_DIRECTORY_NORMAL
    except:
        logger.info("Failed to read Oracle client value")

    return directory


# Check Manhattan based on PO number
def checkPOManhattan(PO):
    cursor_named_params = {'PO_ORDER': PO}

    cursor.execute("""
    select o.PO_ORDER, p.PR_OWN
    from ORDERH o
    inner join PROP p on o.COST_CENTRE = p.PR_SNAM
    where o.PO_ORDER = CAST(TRIM(:PO_ORDER) as NCHAR(20))
    """, cursor_named_params)

    results = cursor.fetchone()

    if results:
        logger.info("PO results = " + str(results))
        directory = output_directory(results[1])

        # logger.info("PO exists in Manhattan")
        return directory
    else:
        logger.info("PO or Property does not exist in Manhattan")
        return False


def move_file(filename):
    outputDirectory = OUTPUT_DIRECTORY_FAILURE
    
    pattern = r"(?i)(\d+)-(sdkt|rca|swms|ecbd|comp).(\w+)"
    match = re.search(pattern, filename)
    if match:
        PO = match.group(1)
        # fileExt = match.group(2)
        outputDirectory = checkPOManhattan(PO)
    else:
        logger.info("Filename does not match pattern")

    shutil.move(os.path.join(INPUT_DIRECTORY, filename), os.path.join(outputDirectory, filename))

    return


def main():

    start = timeit.default_timer()
    
    logger.info("file move start")

    for f in os.listdir(INPUT_DIRECTORY):
        if os.path.isfile(os.path.join(INPUT_DIRECTORY, f)):
            cleaned_name = CleanCharacters(f.encode("ascii", "ignore").decode())
            logger.info("FILE: " + cleaned_name)
            move_file(f)

    stop = timeit.default_timer()

    logger.info("file move runtime: " + "{0:.2f}".format(stop - start) + "s\n")
    
    return True


if __name__ == '__main__': main()

