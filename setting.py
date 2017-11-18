import yaml
import os
PROPERTIES = yaml.load(open('./conf/process.yaml'))
PROJECT_INTER_OUTPUT_DIR = './output'
EXCEL_SUFFIX = '.xls'
if not os.path.exists(PROJECT_INTER_OUTPUT_DIR):
    os.makedirs(PROJECT_INTER_OUTPUT_DIR)
