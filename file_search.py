import os, re, sys
from stat import S_ISREG
from IPython.display import HTML
from textwrap import wrap, indent
from os.path import abspath, join, splitext, split
from os import mkdir, walk, remove
from io import StringIO
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import zipfile
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'

def search_files(pattern, top_dir, extensions=None,
                 exclude_folders=None, size_limit=0,
                 print_errors=True):
    '''
    Search file contents for a regex pattern

    Parameters
    ----------
    pattern : str
        Each line will be treated as an OR search term.
        String can contain regular expressions.
    top_dir : str
        Top directory to perform deep traverse.
    extensions : str, list, default: None
        String or list of file extensions to perform content search.
    exclude_folders : list, default: None
        List of folders to exclude from search.
    size_limit : int, default: 0
    print_errors : bool, default: True
    '''
    file_list = []
    pat = pattern.strip()
    pat = r'(?:{})'.format('|'.join(line.strip() for line in pat.split('\n')))
    print("Searching for regex: {} in {}\n".format(pat, top_dir))
    pat = re.compile(pat, re.IGNORECASE)
    for file_name in files_to_search(top_dir, extensions, exclude_folders, size_limit):
        try:
            if search_file(file_name, pat):
                file_list.append(file_name)
                print('\n      '.join(wrap(file_name, 120)))
        
        except Exception as e:
            if print_errors:
                msg = '{}'.format('<br>'.join(wrap(file_name, 120)))
                display(HTML('<p style="color:red; padding-left:2em; text-indent:-2em">Error reading: {}</p>'.format(msg)))
    return file_list

def files_to_search(top_dir, extensions, exclude_folders, size_limit):
    extensions = extensions or ''
    """yield up full pathname for only files we want to search"""
    for fname in walk_files(top_dir, extensions, exclude_folders):
        try:
            # if it is a regular file and big enough, we want to search it
            sr = os.stat(fname)
            if S_ISREG(sr.st_mode) and sr.st_size >= size_limit:
                yield fname
        except OSError:
            pass

def walk_files(top_dir, extensions, exclude_folders):
    extensions = extensions or ['']
    exclude_folders = exclude_folders or ('')
    extensions = extensions if isinstance(extensions, tuple) else tuple(extensions)
    exclude_folders = exclude_folders if isinstance(exclude_folders, tuple) else tuple(exclude_folders)
    """yield up full pathname for each file in tree under top_dir"""
    for dirpath, dirnames, filenames in os.walk(top_dir, topdown=True):
        dirnames[:] = [d for d in dirnames if d not in exclude_folders]
        filenames = [f for f in filenames if f.endswith(extensions)]
        filenames = [f for f in filenames if not f.startswith('~')]
        
        for fname in filenames:
            pathname = os.path.join(dirpath, fname)
            yield pathname

def search_file(filename, pat):
    #Get plain text from each file and search for pat
    ext = splitext(filename)[1]

    if re.search(pat, filename):
        return True
    
    elif ext == ".pdf":
            content = read_pdf(filename)
            if re.search(pat, content):
                return filename
            
    if ext == '.docx':
        document = zipfile.ZipFile(filename)
        xml_content = document.read('word/document.xml')
        document.close()
        tree = XML(xml_content)

        paragraphs = []
        for paragraph in tree.getiterator(PARA):
            for node in paragraph.getiterator(TEXT):
                if node.text:
                    if re.search(pat, node.text):
                        return True
    elif ext == '.xlsx':
        document = zipfile.ZipFile(filename)
        xml_content = document.read('xl/sharedStrings.xml').decode('utf-8', errors='ignore')
        document.close()
        if re.search(pat, xml_content):
            return True
    else:
        with open(filename, 'rt', encoding="utf8", errors='ignore') as f:
            for line in f:
                if re.search(pat, line):
                    return True

def read_pdf(fname, pages=None):
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)

    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    infile = open(fname, 'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)
        
    infile.close()
    converter.close()
    text = output.getvalue()
    text = text.replace("\xa0", " ")
    output.close
    return text 
