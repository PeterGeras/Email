import os
import logzero
from logzero import logger  # logger is called from other files
import logging

# Folders and files
cwd = os.getcwd()
log_file = os.path.join(cwd, r'logs\main_log.log')

# Setup rotating logfile with 3 rotations, each with a maximum filesize of 1MB:
logzero.logfile(log_file, maxBytes=int(1e6), backupCount=3)

# Set a minimum log level
logzero.loglevel(logging.DEBUG)

# Set a custom formatter
formatter = logging.Formatter('%(asctime)-15s - %(levelname)s: %(message)s')
logzero.formatter(formatter)
