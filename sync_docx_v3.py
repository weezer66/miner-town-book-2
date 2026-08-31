from docx import Document

# Load the DOCX file
doc = Document('manuscript/chapter-01-assigned-to-witness.docx')

# Find paragraph 49
target_para = doc.paragraphs[49]
print(f'Paragraph text: {target_para.text}')
print(f'Number of runs: {len(target_para.runs)}')

# Print each run
for j, run in enumerate(target_para.runs):
    print(f'  Run {j}: "{run.text}"')

# Now replace the content
# The paragraph likely spans multiple runs
if 'The first Initiation queue' in target_para.text:
    print('\nPerforming replacement...')
    
    # Clear all runs
    for run in target_para.runs:
        run.text = ''
    
    # Add the new text
    target_para.text = 'The Initiation queue populated on screen—three entries materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity\'s inventory.'
    
    print('Replacement complete!')

# Save
doc.save('manuscript/chapter-01-assigned-to-witness.docx')
print('DOCX saved')
