import openpyxl
from configuration import retrieve_input_file_path


def load_workbook(file_path=None):
    resolved_file_path = file_path if file_path else retrieve_input_file_path()

    return openpyxl.load_workbook(resolved_file_path, data_only=False)
