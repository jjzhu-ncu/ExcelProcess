import yaml
import os
from logger import LOGGER
PROPERTIES = yaml.load(open('./conf/process.yaml'))
LOGGER.info("project properties: %s" % str(PROPERTIES))
PROJECT_INTER_OUTPUT_DIR = '.\\output'
EXCEL_SUFFIX = '.xls'
if not os.path.exists(PROJECT_INTER_OUTPUT_DIR):
    os.makedirs(PROJECT_INTER_OUTPUT_DIR)
