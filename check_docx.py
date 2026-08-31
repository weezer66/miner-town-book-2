from docx import Document

# Load
docx_path = r'C:\Users\Prisha\OneDrive\Documents\Miner Town\Book 2\manuscript\chapter-01-assigned-to-witness.docx'
doc = Document(docx_path)

# Find and replace
old = 'Then the central grid flickered.\n\nThe Initiation queue populated on screen—the names materialized like component numbers in a system ledger'
new = 'Then the central grid flickered.\n\nThe Initiation queue populated on screen—sixteen names materialized like component numbers in a system ledger'

for para in doc.paragraphs:
    if old in para.text or 'Initiation queue populated on screen' in para.text:
        print(f'Checking: {para.text[:80]}')

print('Update will be done manually - file is locked')
