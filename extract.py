import os, csv, re
import docx

DOCS_PATH = '' # Path to directory in which docx files are

def process_file(root, filename):
    # Article information is contained within several lines. Each article is separated by a blank line.
    # Bold lines should be ignored (except to set current subject) as they contain section information
    # Expected format of article information by line is:
        # Title
        # Journal
        # Issue details then name (after final comma in line)
        # Article first published online date, then DOI (after final comma in line)
        # Link to article online
    doc = docx.Document(os.path.join(root, filename))
    para_iter = iter(doc.paragraphs)
    subject = ''
    for line in para_iter:
        # Ignore bold lines as they are section headings
        if len(line.runs) > 0 and line.runs[0].font.bold:
            if line.runs[0].font.underline:
                subject = line.text
            continue
        # Ignore empty lines and checkboxes as they separate articles from each other
        elif line._p.xpath('string()') != '' and 'PRIVATE "<INPUT TYPE' not in line._p.xpath('string()'):
            article = dict()
            article['Subject'] = subject
            # We use line._p.xpath("string()") rather than the line.text attribute provided by docx, because if the line is a hyperlink, docx will treat it as blank
            article['Title'] = line._p.xpath("string()")
            if 'HYPERLINK' in article['Title']: # This is a sign of an error
                article['Title'] = ''
            line = next(para_iter)
            article['Journal'] = line._p.xpath("string()").title()
            if not any(x == article['Journal'] for x in ['Nations And Nationalism', 'Studies In Ethnicity And Nationalism']): # This is a sign of an error
                article['Journal'] = ''
            line = next(para_iter)
            line_split = re.split('(\d,)', line._p.xpath("string()"))
            article['Volume'] = ''.join(line_split[0:-1]).rstrip(',')
            article['Author'] = line_split[-1].strip().title().replace(' And ', ' and ')
            line = next(para_iter)
            article['Published'] = line._p.xpath("string()").replace('Article first published online : ', '').replace('Version of Record online : ', '')
            line = next(para_iter)
            article['Link'] = line._p.xpath("string()")
            yield article

if __name__ == '__main__':
    with open('articles.csv', 'w') as csvfile:
        # fieldnames = ['Subject', 'Title', 'Journal', 'Volume', 'Author', 'Published']
        book_reviews_strs = ['book reviews', 'books reviews']
        fieldnames = ['Parse errors', 'Subject', 'Title', 'Journal', 'Volume', 'Author', 'Published', 'Link']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        error_writer = csv.writer(csvfile)
        writer.writeheader()
        # Get all docx files in directory
        for root, dirs, files in os.walk(DOCS_PATH):
            # Process each file
            for filename in sorted(files):
                # Only process docx files found in directory
                if re.match('.+\.docx$', filename):
                    # Iterate over all articles in file (an article is a collection of lines in the file, separated from other articles by blank lines)
                    iter_articles = process_file(root, filename)
                    for article in iter_articles:
                        error_fields = list()
                        for key, val in article.items():
                            # Add errors for empty fields, unless the empty field is 'Author' and this is a 'book reviews' article (as there isn't an author here)
                            if val == '' and not (key == 'Author' and any(x == article['Title'].lower() for x in book_reviews_strs)):
                                error_fields.append(key)
                        if len(error_fields) > 0:
                            article['Parse errors'] = 'Errors / missing: %s' % (', '.join(error_fields),)
                        else:
                            article['Parse errors'] = ''
                        print(article)
                        writer.writerow(article)
                        print('Wrote article to csv\n')