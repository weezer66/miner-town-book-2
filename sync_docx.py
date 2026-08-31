from docx import Document

# Load the DOCX file
doc = Document('manuscript/chapter-01-assigned-to-witness.docx')

# The old and new passages
old_text = 'Then the central grid flickered. The first Initiation queue populated in red. Three entries. Elara leaned toward the screen. The names belonged to Miner Town.'
new_text = 'Then the central grid flickered.\n\nThe Initiation queue populated on screen—three entries materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity\'s inventory.\n\nElara leaned toward the screen. The names belonged to Miner Town.'

# Search and replace in all paragraphs
found = False
for paragraph in doc.paragraphs:
    if old_text in paragraph.text:
        # Rebuild the paragraph with new text
        paragraph.clear()
        for line in new_text.split('\n'):
            if line.strip():
                p = paragraph.add_run(line)
                if paragraph.style.font.size:
                    p.font.size = paragraph.style.font.size
        found = True
        print('Updated passage in DOCX')

if not found:
    # Try searching for partial match
    for paragraph in doc.paragraphs:
        if 'The first Initiation queue' in paragraph.text:
            print(f'Found but not exact match: {paragraph.text[:100]}')

# Save the updated DOCX
doc.save('manuscript/chapter-01-assigned-to-witness.docx')
print('DOCX file refreshed successfully')
