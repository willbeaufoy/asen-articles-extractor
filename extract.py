import os, csv, re
import docx

DOCS_PATH = '/home/sofr/d/current subject themes/current subject themes' # Path to directory in which docx files are
# DOCS_PATH = '/home/sofr/d/test asen'

def process_file(root, filename):
    # Article information is contained within several lines. Each article is separated by a blank line.
    # Bold lines should be ignored as they contain section information
    # Expected format of article information by line is:
        # Title
        # Journal
        # Issue details then name (after final comma in line)
        # Article first published online date, then DOI (after final comma in line)
        # Link to article online
    doc = docx.Document(os.path.join(root, filename))
    para_iter = iter(doc.paragraphs)
    doc_end = False
    subject = ''
    for line in para_iter:
        # Ignore bold lines as they are section headings, and empty lines as they separate articles from each other
        if len(line.runs) > 0 and line.runs[0].font.bold:
            if line.runs[0].font.underline:
                subject = line.text
            continue
        elif line.text != '' or line._p.xpath("string()") != '':
            # article = Article(filename.rstrip('.docx'))
            article = dict()
            article['Subject'] = subject
            # article.title = line.text
            if line.text != '':
                article['Title'] = line.text
            else:
                article['Title'] = line._p.xpath("string()")
            line = next(para_iter)
            # article.journal = line.text.title()
            if line.text != '':
                article['Journal'] = line.text.title()
            else:
                article['Journal'] = line._p.xpath("string()").title()
            line = next(para_iter)
            if line.text != '':
                line_split = re.split('(\d,)', line.text)
            else:
                line_split = re.split('(\d,)', line._p.xpath("string()"))
            # article.volume = line[0].strip()
            article['Volume'] = ''.join(line_split[0:-1]).rstrip(',')
            # article.author = line[1].strip().title()
            article['Author'] = line_split[-1].strip().title().replace('And', 'and')
            line = next(para_iter)
            # article.published = line.text.replace('Article first published online : ', '').replace('Version of Record online : ', '')
            if line.text != '':
                article['Published'] = line.text.replace('Article first published online : ', '').replace('Version of Record online : ', '')
            else:
                article['Published'] = line._p.xpath("string()").replace('Article first published online : ', '').replace('Version of Record online : ', '')
            line = next(para_iter)
            # article['Link'] = line._p.getchildren()[0].getchildren()[0].getchildren()[1].text
            article['Link'] = line._p.xpath("string()")
            yield article

        # except StopIteration:
        #     doc_end = True
        #     continue

if __name__ == '__main__':
    with open('articles.csv', 'w') as csvfile:
        # fieldnames = ['Subject', 'Title', 'Journal', 'Volume', 'Author', 'Published']
        fieldnames = ['Subject', 'Title', 'Journal', 'Volume', 'Author', 'Published', 'Link']
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
                        # print(article, '\n')
                        article_ok = True
                        for val in article.values():
                            if val == '':
                                article_ok = False
                                break
                        if article_ok:
                            writer.writerow(article)
                        else:
                            error_writer.writerow(['ERROR parsing article in %s' % (filename)])
