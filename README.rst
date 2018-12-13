.. -*- mode: rst -*-

file_search
============

Search for text strings within file contents in a given directory.

Pattern to search for can include regex. This function can search within
``.doc``, ``.docx``, ``.xls``, ``.xlsx`` and ``.pdf`` files, as well as
regular text-like files, such as ``.ipynb``, ``.txt`` or ``.csv``


Dependencies
~~~~~~~~~~~~

file_search requires:

- Python (>= 3.4)
- pdfminer
- zipfile


Usage
~~~~~~~~~~~~

Designed to be used within a Jupyter Notebook::

    from file_search import search_files

    search_directory = r'C:\Users\{user}\Documents'
    exclude_folders = []
    extensions = ['.xlsx', '.xls', '.docx']
    size_limit = 100

    patterns = '''
    string1
    string1
    string1
    '''

    results = search_files(pattern=patterns,
                           top_dir=search_directory,
                           extensions=extensions,
                           exclude_folders=exclude_folders,
                           size_limit=size_limit,
                           print_errors=True)