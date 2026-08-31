from docx import Document
from docx.shared import Pt

# Load the DOCX file
doc = Document('manuscript/chapter-01-assigned-to-witness.docx')

# Search for the paragraph containing the old text
target_phrase = 'The first Initiation queue'
new_passage = 'The Initiation queue populated on screen—three entries materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity\'s inventory.'

found_and_replaced = False

for i, paragraph in enumerate(doc.paragraphs):
    # Look for the paragraph with the old text
    if target_phrase in paragraph.text:
        print(f'Found target in paragraph {i}: {paragraph.text[:80]}...')
        
        # Get the paragraph text for reference
        para_text = paragraph.text
        
        # Replace the specific phrase
        if 'The first Initiation queue populated in red. Three entries.' in para_text:
            # Clear the paragraph and rebuild
            paragraph.clear()
            paragraph.add_run('Then the central grid flickered.\n\n')
            paragraph.add_run(new_passage)
            paragraph.add_run('\n\nElara leaned toward the screen. The names belonged to Miner Town.')
            found_and_replaced = True
            print('Successfully updated the passage!')
            break

if not found_and_replaced:
    print('Exact phrase not found. Trying alternative approach...')
    for i, paragraph in enumerate(doc.paragraphs):
        if 'Initiation queue' in paragraph.text and 'three entries' in paragraph.text.lower():
            print(f'Found at paragraph {i}')
            print(f'Current text: {paragraph.text}')

# Save
doc.save('manuscript/chapter-01-assigned-to-witness.docx')
print('DOCX file saved')
