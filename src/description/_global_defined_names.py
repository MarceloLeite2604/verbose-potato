import json
from workbook import load_workbook
from configuration import retrieve_output_file_path
from util.cell import split_cell_reference


def _retrieve_type(defined_name_range):

    if not defined_name_range:
        return None

    if 'end' in defined_name_range:
        if defined_name_range['start']['col'] == defined_name_range['end']['col'] \
                or defined_name_range['start']['row'] == defined_name_range['end']['row']:
            return 'list'
        else:
            return 'table'

    return 'constant'


def _retrieve_cell_reference(defined_name):

    (worksheet_title, cell_reference_text) = next(defined_name.destinations)

    if not cell_reference_text:
        return None

    cell_reference = split_cell_reference(cell_reference_text)

    if not cell_reference:
        raise ValueError(
            f'Unknown cell reference format \"{cell_reference_text}\" for defined name \"{defined_name.name}\".')

    data = {
        'worksheet_title': worksheet_title
    }

    if 'start' in cell_reference:
        data['start'] = cell_reference['start']

    if 'end' in cell_reference:
        data['end'] = cell_reference['end']

    return data


def _elaborate_coordinate(cell_reference):
    coordinate = f"{cell_reference['worksheet_title']}!"
    if 'end' in cell_reference:
        coordinate += f"{cell_reference['start']['col']}{cell_reference['start']['row']}:"
        coordinate += f"{cell_reference['end']['col']}{cell_reference['end']['row']}"
    else:
        coordinate += f"{cell_reference['start']['col']}{cell_reference['start']['row']}"
    return coordinate


def _elaborate_text(title, data):

    coordinate = _elaborate_coordinate(data['cell_reference'])

    return f"Defined name \"{title}\" is a {data['type']} defined at \"{coordinate}\""


def _elaborate_metadata(title, data):
    return {
        'name': title,
        'type': data['type'],
        'coordinate': _elaborate_coordinate(data['cell_reference'])
    }


def _elaborate_data(defined_name_title, defined_name):
    defined_name_cell_reference = _retrieve_cell_reference(defined_name)
    defined_name_type = _retrieve_type(defined_name_cell_reference)

    if not defined_name_type or not defined_name_cell_reference:
        return None

    defined_name_data = {
        'type': defined_name_type,
        'cell_reference': defined_name_cell_reference
    }

    output = {}

    output['text'] = _elaborate_text(
        defined_name_title, defined_name_data)

    output['metadata'] = _elaborate_metadata(
        defined_name_title, defined_name_data)

    return output


def _retrieve_definitions():
    workbook = load_workbook()

    definitions = {
        defined_name: _elaborate_data(
            defined_name,
            workbook.defined_names[defined_name])
        for defined_name in workbook.defined_names}

    # Remove keys with None values
    definitions = {k: v for k, v in definitions.items() if v}

    return definitions


def write_definition():

    definitions = _retrieve_definitions()

    output_path = retrieve_output_file_path(
        'definitions', 'global-definitions.json')

    with open(output_path, 'w') as file:
        json.dump(definitions, file, ensure_ascii=False,
                  indent=2, separators=(',', ': '))
