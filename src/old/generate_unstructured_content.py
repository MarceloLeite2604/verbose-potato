import os
from dotenv import load_dotenv, find_dotenv

from unstructured.partition.xlsx import partition_xlsx
from unstructured.staging.base import convert_to_dict

load_dotenv(find_dotenv(), override=True)

spreadsheet_location = os.environ['INPUT_FILE_PATH']

elements = partition_xlsx(filename=spreadsheet_location,
                          chunking_strategy='by_title')

isd = convert_to_dict(elements)

print(isd)
