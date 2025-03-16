from itertools import islice
import json
import os
from pinecone import Pinecone, ServerlessSpec
from langchain_openai import OpenAIEmbeddings
from hashlib import sha256

from configuration import retrieve_output_file_path

_PINECONE: Pinecone | None = None
_EMBEDDINGS: OpenAIEmbeddings | None = None

_INDEX_NAME = "verbose-potato"


def _create_index_if_nonexistent():

    if _INDEX_NAME not in _PINECONE.list_indexes().names():
        print(f'Creating Pinecone index \"{_INDEX_NAME}\".')

        _PINECONE.create_index(
            name=_INDEX_NAME,
            dimension=1536,
            metric='cosine',
            spec=ServerlessSpec(
                cloud='aws',
                region='us-east-1'
            )
        )


def _append_vectors(documents):

    print('Appeding vectors.')

    vectors_output_file = retrieve_output_file_path(
        'descriptions', 'vectors.json')

    if os.path.exists(vectors_output_file):
        print('Vectors cache found.')
        with open(vectors_output_file) as file:
            cache = json.load(file)
    else:
        cache = {}

    texts_to_embed = []
    texts_to_embed_ids = []

    for document in documents.values():
        document_id = document['id']
        if document_id not in cache:
            texts_to_embed.append(document['text'])
            texts_to_embed_ids.append(document_id)
        else:
            document['vector'] = cache[document_id]

    if len(texts_to_embed) > 0:
        print(f'{len(texts_to_embed)} texts to embed.')

        global _EMBEDDINGS
        if not _EMBEDDINGS:
            print('Creating OpenAI Embeddings.')
            _EMBEDDINGS = OpenAIEmbeddings(model="text-embedding-ada-002")

        print('Embedding documents.')
        vectors = _EMBEDDINGS.embed_documents(texts_to_embed)
        print('Embedding complete.')

        for [document_id, vector] in zip(texts_to_embed_ids, vectors):
            cache[document_id] = vector

        with open(vectors_output_file, 'w') as file:
            json.dump(cache, file)

    for document in documents.values():
        if 'vector' not in document:
            document['vector'] = cache[document['id']]

    return documents


def _elaborate_vectors(descriptions):

    for description in descriptions.values():
        encoded_text = description['text'].encode()
        description['id'] = sha256(encoded_text).hexdigest()

    updated_descriptions = _append_vectors(descriptions)

    vectors = []

    for description in updated_descriptions.values():

        vectors.append({
            'id': description['id'],
            'values': description['vector'],
            'metadata': description['metadata']
        })

    return vectors


def upsert(namespace: str, descriptions):

    # sorted_descriptions = dict(sorted(descriptions.items()))

    # Create a new dictionary with only the 10 first items of sorted_descirpitons
    # sliced_descriptions = dict(islice(sorted_descriptions.items(), 10))
    sliced_descriptions = descriptions

    global _PINECONE
    if not _PINECONE:
        print('Creating Pinecone client.')
        _PINECONE = Pinecone()

    _create_index_if_nonexistent()

    print('Elaborating vectors.')
    vectors = _elaborate_vectors(sliced_descriptions)

    print('Upserting vectors.')
    _PINECONE.Index(_INDEX_NAME).upsert(
        namespace=namespace,
        vectors=vectors,
        batch_size=100,
        show_progress=True)
