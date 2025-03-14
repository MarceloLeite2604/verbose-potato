import os
from pathlib import Path
from dotenv import load_dotenv, find_dotenv
import tempfile

_OUTPUT_DIRECTORY = Path('output-files')

env_loaded = False


def _load_env():
    global env_loaded
    if not env_loaded:
        load_dotenv(find_dotenv(), override=True)
        env_loaded = True


def retrieve_input_file_path():
    _load_env()
    return os.environ['INPUT_FILE_PATH']


def retrieve_output_file_path(*paths):

    output_path = _OUTPUT_DIRECTORY.joinpath(*paths)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    return output_path.as_posix()


def create_temporary_file(suffix: str | None = None):

    _, file_path = tempfile.mkstemp(suffix=suffix)

    return file_path
