"""Test the base module."""


def test_msword_demo():
    # msword demo
    # Basic Demo
    from msword import (
        LocalDocxStore,  # Local files store returning, as values, docx.document.Document objects.
        LocalDocxTextStore,  # Local files store returning, as values, text extracted from the documents.
        AllLocalFilesDocxStore,  # Like LocalDocxStore but doesn't filter for valid msword extensions.
        AllLocalFilesDocxTextStore,  # Like LocalDocxTextStore but doesn't filter for valid msword extensions.
        # ----------------------------------------------------------------------------------
        # Mapping wrappers (codecs)
        only_files_with_msword_extension,  # Wrap a Mapping to filter for valid msword extensions (.doc and .docx).
        with_bytes_to_text_decoding,  # Wrap a Mapping to decode bytes to text extracted from docx.Document objects.
        with_bytes_to_doc_decoding,  # Wrap a Mapping to decode bytes to docx.Document objects.
        with_doc_to_text_decoding,  # Wrap a Mapping to decode docx.Document objects to text.
    )
    import docx
    import pytest
    import os

    # Import a utility that provides a directory with test data
    from msword.tests.util import test_data_dir

    # --------------------------------------------------------------------------
    # A store of local MS Word files
    # Create a store that returns text extracted from docx files.
    docs_text_content = LocalDocxTextStore(test_data_dir)
    # List the keys of the store (should be relative paths of MS Word files)
    _ = sorted(docs_text_content)

    # Access the text content of a specific document.
    simple_doc_text = docs_text_content["simple.docx"]

    # --------------------------------------------------------------------------
    # Deconstructing LocalDocxTextStore:
    # The base store provides all files: keys are relative paths, and values are bytes.
    from dol import Files

    raw_files = Files(test_data_dir)
    _ = sorted(raw_files)
    # Get the raw bytes of 'simple.docx'
    b = raw_files["simple.docx"]
    assert isinstance(b, bytes), "Value should be bytes"
    # For debugging purposes, you can print b or inspect it:
    # print(b)

    # --------------------------------------------------------------------------
    # with_bytes_to_doc_decoding
    # Convert byte-valued mappings to return docx.Document objects.
    doc_objects = with_bytes_to_doc_decoding(raw_files)
    _ = sorted(doc_objects)
    doc = doc_objects["simple.docx"]
    assert isinstance(doc, docx.document.Document), "Value should be a Document object"
    # Verify that the document has the expected number of paragraphs.
    assert len(doc.paragraphs) == 4
    # Aggregate the text of all paragraphs in the document.
    doc_paragraphs_text = "\n".join(p.text for p in doc.paragraphs)
    assert (
        doc_paragraphs_text
        == "Just a bit of text to show that is works. Another sentence.\nThis is after a newline.\n\nThis is after two newlines."
    )

    # --------------------------------------------------------------------------
    # with_bytes_to_text_decoding
    # If only aggregated text is needed, convert byte-valued mappings to return text.
    doc_texts = with_bytes_to_text_decoding(raw_files)
    _ = sorted(doc_texts)
    assert isinstance(doc_texts["simple.docx"], str), "Value should be a string"
    simple_doc_text_again = doc_texts["simple.docx"]
    assert (
        simple_doc_text_again
        == "Just a bit of text to show that is works. Another sentence.\nThis is after a newline.\n\nThis is after two newlines."
    )

    # --------------------------------------------------------------------------
    # Filter out non-MS Word files
    # Although doc_texts lists all files (except hidden ones) of all extensions,
    # accessing a file that is not a valid MS Word document should raise an exception.
    with pytest.raises(Exception):
        # Expect an exception when trying to decode a non-MS Word file.
        _ = doc_texts["not_an_msword_doc.txt"]

    # If you want to only see files with MS Word extensions ('.doc' and '.docx'),
    # wrap your store with only_files_with_msword_extension.
    with_extension_filter = only_files_with_msword_extension(doc_texts)
    assert sorted(with_extension_filter) == ["simple.docx", "with_doc_extension.doc"]

    # --------------------------------------------------------------------------
    # Sourcing your MS Word content from anywhere
    # MS Word content might be sourced from remote storage, databases, or zip files.
    # Here we demonstrate sourcing from a zip file.
    from dol import FilesOfZip, Pipe

    # Create a pipeline that starts with a zip store maker, then filters and decodes.
    msword_doc_texts_of_zip = Pipe(
        FilesOfZip,
        only_files_with_msword_extension,  # Filter out non-MS Word extensions.
        with_bytes_to_text_decoding,  # Return values as text.
    )

    # Sourcing the zip: you can pass a zip file path or zip bytes.
    zip_file_path = os.path.join(test_data_dir, "some_zip_file.zip")
    zipped_doc_texts = msword_doc_texts_of_zip(zip_file_path)

    # List the keys in the zipped store.
    assert sorted(zipped_doc_texts) == ["simple.docx", "with_doc_extension.doc"]

    # Access the text content of 'simple.docx' from the zipped store.
    assert zipped_doc_texts["simple.docx"] == doc_texts["simple.docx"]

    # --------------------------------------------------------------------------
    # Removing extensions from keys
    # Use KeyCodecs to map keys, e.g., remove file extensions.
    from dol import KeyCodecs

    remove_extensions = KeyCodecs.mapped_keys(
        zipped_doc_texts, decoder=lambda x: os.path.splitext(x)[0]
    )
    zipped_doc_texts_without_extensions = remove_extensions(zipped_doc_texts)
    expected_keys = ["simple", "with_doc_extension"]
    assert sorted(zipped_doc_texts_without_extensions) == expected_keys
    # Ensure that the text for key 'simple' matches that for 'simple.docx'
    assert (
        zipped_doc_texts_without_extensions["simple"] == zipped_doc_texts["simple.docx"]
    )
