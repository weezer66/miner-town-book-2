from docx import Document

# Use absolute path with Windows backslashes
docx_path = r'C:\Users\Prisha\OneDrive\Documents\Miner Town\Book 2\manuscript\chapter-01-assigned-to-witness.docx'

# Load the DOCX file
doc = Document(docx_path)

# Find and replace the specific paragraph
old_phrase = 'The Initiation queue populated on screen—three entries materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity\'s inventory.'
new_phrase = 'The Initiation queue populated on screen—the names materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity\'s inventory.'

for para in doc.paragraphs:
    if old_phrase in para.text:
        para.text = new_phrase
        print(f'Updated: {para.text[:80]}...')

# Save
doc.save(docx_path)
print('DOCX refreshed')
