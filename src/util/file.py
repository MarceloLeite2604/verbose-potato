from unidecode import unidecode


def as_file_name(value: str):
    return unidecode(value.replace(' ', '-').lower())
