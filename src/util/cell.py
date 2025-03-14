
import regex


_CELL_REFERENCE_REGEX = r"^(\$?'?(?<workbook_name>[A-zÀ-ú ]+)'?[!\.])?\$?(?<start_col>[A-Z]+)\$?(?<start_row>\d+)(?::\$?(?<end_col>[A-Z]+)\$?(?<end_row>\d+))?$"


def split_cell_reference(coordinate):
    regex.compile(_CELL_REFERENCE_REGEX)
    match = regex.match(_CELL_REFERENCE_REGEX, coordinate)

    if not match:
        return None

    data = {
        'workbook_name': match['workbook_name'],
        'start': {
            'row': match['start_row'],
            'col': match['start_col']
        }
    }

    if match['end_row'] and match['end_col']:
        data['end'] = {
            'row': match['end_row'],
            'col': match['end_col']
        }

    return data
