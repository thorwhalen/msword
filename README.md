
# msword
Simple mapping view to docx (Word Doc) elements


To install:	```pip install msword```


# Examples

## LocalDocxTextStore
Local files store returning, as values, text extracted from the documents.
Use this when you just want the text contents of the document.
If you want more, you'll need to user LocalDocxStore with the appropriate content extractor
(i.e. the obj_of_data function in a py2store.wrap_kvs wrapper).

Note: Filters for valid msword extensions (.doc and .docx).
To NOT filter for valid extensions, use ``AllLocalFilesDocxTextStore`` instead.

```python
>>> from msword import LocalDocxTextStore, test_data_dir
>>> import docx
>>> s = LocalDocxTextStore(test_data_dir)
>>> assert {'more_involved.docx', 'simple.docx'}.issubset(s)
>>> v = s['simple.docx']
>>> assert isinstance(v, str)
>>> print(v)
Just a bit of text to show that is works. Another sentence.
This is after a newline.
<BLANKLINE>
This is after two newlines.
```

## LocalDocxStore

Local files store returning, as values, docx objects.
Note: Filters for valid msword extensions (.doc and .docx).
To Note filter for valid extensions, use ``AllLocalFilesDocxStore`` instead.

```python
>>> from msword import LocalDocxStore, test_data_dir
>>> import docx
>>> s = LocalDocxStore(test_data_dir)
>>> assert {'more_involved.docx', 'simple.docx'}.issubset(s)
>>> v = s['more_involved.docx']
>>> assert isinstance(v, docx.document.Document)
```

What does a ``docx.document.Document`` have to offer?
If you really want to get into it, see here: https://python-docx.readthedocs.io/en/latest/

Meanwhile, we'll give a few examples here as an amuse-bouche.

```python
>>> ddir = lambda x: set([xx for xx in dir(x) if not xx.startswith('_')])  # to see what an object has
>>> assert ddir(v).issuperset({
...     'add_heading', 'add_page_break', 'add_paragraph', 'add_picture', 'add_section', 'add_table',
...     'core_properties', 'element', 'inline_shapes', 'paragraphs', 'part',
...     'save', 'sections', 'settings', 'styles', 'tables'
... })
```

``paragraphs`` is where the main content is, so let's have a look at what it has.

```python
>>> len(v.paragraphs)
21
>>> paragraph = v.paragraphs[0]
>>> assert ddir(paragraph).issuperset({
...     'add_run', 'alignment', 'clear', 'insert_paragraph_before',
...     'paragraph_format', 'part', 'runs', 'style', 'text'
... })
>>> paragraph.text
'Section 1'
>>> assert ddir(paragraph.style).issuperset({
...     'base_style', 'builtin', 'delete', 'element', 'font', 'hidden', 'locked', 'name', 'next_paragraph_style',
...     'paragraph_format', 'part', 'priority', 'quick_style', 'style_id', 'type', 'unhide_when_used'
... })
>>> paragraph.style.style_id
'Heading1'
>>> paragraph.style.font.color.rgb
RGBColor(0x2f, 0x54, 0x96)
```

You get the point...

If you're only interested in one particular aspect of the documents, you should your favorite
py2store wrappers to get the store you really want. For example:

```python
>>> from py2store import wrap_kvs
>>> ss = wrap_kvs(s, obj_of_data=lambda doc: [paragraph.style.style_id for paragraph in doc.paragraphs])
>>> assert ss['more_involved.docx'] == [
...     'Heading1', 'Normal', 'Normal', 'Heading2', 'Normal', 'Normal',
...     'Heading1', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal', 'Normal',
...     'ListParagraph', 'ListParagraph', 'Normal', 'Normal', 'ListParagraph', 'ListParagraph', 'Normal'
... ]
```

The most common use case is probably getting text, not styles, out of a document.
It's so common, that we've done the wrapping for you:
Just use the already wrapped LocalDocxTextStore store for that purpose.


