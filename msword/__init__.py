"""Simple access to docx (Word Doc) elements."""
from msword.base import (
    LocalDocxStore,  # Local files store returning, as values, ``docx.document.Document`` objects.
    LocalDocxTextStore,  # Local files store returning, as values, text extracted from the documents.
    AllLocalFilesDocxStore,  # Like LocalDocxStore but doesn't filter for valid msword extensions.
    AllLocalFilesDocxTextStore,  # Like LocalDocxTextStore but doesn't filter for valid msword extensions.
    # ----------------------------------------------------------------------------------
    # Mapping wrappers (codecs)
    only_files_with_msword_extension,  # Wrap a Mapping to filter for valid msword extensions (.doc and .docx).
    with_bytes_to_text_decoding,  # Wrap a Mapping to decode bytes to text extracted from docx.Document objects
    with_bytes_to_doc_decoding,  # Wrap a Mapping to decode bytes to docx.Document objects.
    with_doc_to_text_decoding,  # Wrap a Mapping to decode docx.Document objects to text.
)
from msword.tests.util import test_data_dir, test_data_posix
