import os
from util.file import as_file_name
from workbook import load_workbook
from configuration import create_temporary_file, retrieve_output_file_path
from xlsx2html import xlsx2html
from redis import Redis
import json


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

    # xlsx2html(input_file_path, output_file_path,
    #           locale='pt_BR', sheet=worksheet_index)

    metadata_file_path = retrieve_output_file_path(
        'structure', 'metadata.json')

    if os.path.exists(metadata_file_path):
        with open(metadata_file_path, 'r') as metadata_file:
            metadata = json.load(metadata_file)
    else:
        metadata = {}

    metadata[worksheet_title] = file_name

    with open(metadata_file_path, 'w') as metadata_file:
        json.dump(metadata, metadata_file, ensure_ascii=False,
                  indent=2, separators=(',', ': '))


def _write_workbook_structure(input_file_path: str):

    print('Writing workbook structure.')

    workbook = load_workbook(input_file_path)

    for worksheet_index, worksheet_title in enumerate(workbook.sheetnames):
        _write_worksheet_structure(
            input_file_path, worksheet_title, worksheet_index)


def save_structure_on_database():
    # Connect on redis
    password = os.environ['REDIS_PASSWORD']
    redis_client = Redis(
        host='localhost',
        password=password,
        port=6379,
        db=0)

    metadata_file_path = retrieve_output_file_path(
        'structure', 'metadata.json')

    metadata = {}

    with open(metadata_file_path, 'r') as metadata_file:
        metadata = json.load(metadata_file)

    for worksheet_name, structure_file_name in metadata.items():

        print(f'Saving worksheet \"{worksheet_name}\" structure on database.')
        structure_file_path = retrieve_output_file_path(
            'structure', structure_file_name)

        with open(structure_file_path, 'r') as metadata_file:
            structure_content = metadata_file.read()
            # TODO Beware with special characters
            redis_client.set(worksheet_name, structure_content)

    redis_client.close()


def write_structure():

    customized_input = _customize_input()

    _write_workbook_structure(customized_input)

    save_structure_on_database()
