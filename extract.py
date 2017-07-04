DOCS_PATH = '/home/sofr/d/current subject themes/current subject themes'

# Get all docx files in directory
for file in files:
    # Article information is contained within several lines. Each article is separated by a blank line.
    # Bold lines should be ignored as they contain section information
    # Expected format of article information by line is:
        # Title
        # Journal
        # Issue details then name (after final comma in line)
        # Article first published online date, then DOI (after final comma in line)
        # Link to article online