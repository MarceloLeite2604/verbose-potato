from . import _global_defined_names, _worksheets


def write_definitions():

    print('Writing workbook definitions.')
    # _global_defined_names.write_definition()
    _worksheets.write_definitions()
