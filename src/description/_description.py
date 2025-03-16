from . import _global_defined_names
from . import _worksheets


def write_definitions():

    print('Writing workbook definitions.')
    _global_defined_names.save_descriptions()
    # _worksheets.write_definitions()
