from . import _global_defined_names, worksheets


def write_definitions():

    print('Writing workbook definitions.')
    _global_defined_names.write_definition()
    # worksheets.write_definitions()
