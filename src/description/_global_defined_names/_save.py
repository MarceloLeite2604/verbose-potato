import json
from configuration import retrieve_output_file_path
from util.pinecone import upsert

_NAMESPACE = "descriptions"


def save_descriptions():
    input_path = retrieve_output_file_path(
        'descriptions', 'global-definitions.json')

    with open(input_path) as file:
        descriptions = json.load(file)

    upsert(_NAMESPACE, descriptions)
