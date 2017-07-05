# ASEN articles extractor

Extracts data about articles in ASEN journals from docx files into a csv file. Built to work only with this specific word doc format. Do not run on other documents without modifiying code first.

Usage
---------

Ensure the `DOCS_PATH` constant in `extract.py` points to the directory where the docx files to be processed are. Then run:

```python
python3 extract.py
```

The csv file will be created in the directory the module is run from.

Errors
----------

If an article field is not found, this will be indicated in the first csv column 'Parse errors'. The only exception to this is no error will be recorded for a missing 'Author' for book reviews articles, as they don't have authors. Rows with an error column must be manually checked.

Tested with
----------------

- Linux (Ubuntu 17.04)
- Python 3.5
- For required python modules see requirements.txt. To install these run pip install -r requirements.txt