from util.file import as_file_name
from workbook import load_workbook
from configuration import create_temporary_file, retrieve_output_file_path
from xlsx2html import xlsx2html


def _customize_input():

    print('Customizing workbook.')

    workbook = load_workbook()

    customized_input = create_temporary_file('.xlsx')

    for worksheet in workbook.worksheets:

        for row_index, _ in enumerate(worksheet.iter_rows()):
            row_dimentsions = worksheet.row_dimensions[row_index]
            if row_dimentsions.outline_level > 0:
                row_dimentsions.hidden = False

    workbook.save(customized_input)

    return customized_input


def _write_worksheet_structure(
        input_file_path: str,
        worksheet_title: str,
        worksheet_index: int):

    print(f'Writing structure for \"{worksheet_title}\" worksheet.')

    file_name = f'{as_file_name(worksheet_title)}.html'

    output_file_path = retrieve_output_file_path('structure', file_name)

    xlsx2html(input_file_path, output_file_path,
              locale='pt_BR', sheet=worksheet_index)


def _write_workbook_structure(input_file_path: str):

    print('Writing workbook structure.')

    workbook = load_workbook(input_file_path)

    for worksheet_index, worksheet_title in enumerate(workbook.sheetnames):
        _write_worksheet_structure(
            input_file_path, worksheet_title, worksheet_index)


def write_structure():

    customized_input = _customize_input()

    _write_workbook_structure(customized_input)
