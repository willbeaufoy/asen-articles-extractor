# ASEN articles extractor

Extracts data about articles in ASEN journals from docx files into a csv file. Built to work only with this specific word doc format. Do not run on other documents without modifiying code first.

If not all fields are found for an article, module will write 'ERROR parsing article in [filename].docx' into csv file in place of article data. However some articles may have all data filled in, but in the wrong fields. These will need to be manually checked.

Usage
---------

- Ensure the DOCS_PATH constant in `extract.py` points to the directory where the docx files to be processed are.
- Then run:

```python
python3 extract.py
```

- The csv file will be created in the directory the module is run from

Tested with
----------------

- Linux (Ubuntu 17.04)
- Python 3.5
- For required python modules see requirements.txt. To install these run pip install -r requirements.txt