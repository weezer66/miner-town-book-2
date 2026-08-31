from docx import Document

docx_path = r'C:\Users\Prisha\OneDrive\Documents\Miner Town\Book 2\manuscript\chapter-01-assigned-to-witness.docx'

# Load the DOCX
doc = Document(docx_path)

# The old passage to find and replace
old_passage = """Then the central grid flickered.

The Initiation queue populated on screen—sixteen names materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity's inventory.

Elara leaned toward the screen. The names belonged to Miner Town.

And beneath each name, where the system normally displayed the responsible gana, a fifth designation appeared instead.

UNASSIGNED.

UNASSIGNED.

UNASSIGNED.

For the first time that morning, the machine had produced a category Trinity did not recognize.

The red queue continued to populate. Thirty-seven names. Forty-one. Fifty-three.

No one moved. The four monitors stood frozen, watching the display as if the queue might resolve itself if they waited long enough. In Trinity, problems that didn't fit the categories were usually managed by not naming them at all. The system preferred clean labels to inconvenient facts. But this one refused to be ignored."""

# The new expanded passage
new_passage = """Then the central grid flickered.

The Initiation queue populated on screen—sixteen names materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity's inventory. Sixteen children. Four for each monitor. Four for each gana. The mathematics was clean. Predictable.

Elara leaned toward the screen. The names belonged to Miner Town.

The system began its assignment sequence. Four names scrolled to Rowan's EVENT stream. Four to Daren's THRESHOLD. Four to Rohan's CORRECTION protocols. Four to Elara's VARIANCE analysis. Each batch connected to its designated gana with the mechanical precision Trinity preferred. The grid moved through the assignments with invisible efficiency.

Then the central grid flickered again.

Three names remained on the primary display, orphaned. And beneath each name, where the system normally displayed the responsible gana, a fifth designation appeared instead.

UNASSIGNED.

UNASSIGNED.

UNASSIGNED.

A red alert bar materialized beneath the three names. OPERATOR INTERVENTION REQUIRED. The system's guardrails had triggered. Whatever metric or pattern these three children violated, whatever gap in Trinity's control framework they represented, the machine could not proceed without human authorization.

For the first time that morning, the system had encountered something it could not sort.

No one moved. The four monitors stood frozen, watching the display as if the queue might resolve itself if they waited long enough. In Trinity, problems that didn't fit the categories were usually managed by not naming them at all. The system preferred clean labels to inconvenient facts. But this one refused to be ignored."""

# Find and replace
found = False
for i, para in enumerate(doc.paragraphs):
    if 'Then the central grid flickered' in para.text and 'sixteen names' in para.text:
        print(f'Found target paragraph at index {i}')
        found = True
        
        # We need to replace multiple paragraphs. Let's identify the range
        # Start from this paragraph and go until we find the end marker
        start_idx = i
        end_idx = i
        
        # Find the end of the old passage
        for j in range(i, min(i+30, len(doc.paragraphs))):
            if 'But this one refused to be ignored' in doc.paragraphs[j].text:
                end_idx = j
                print(f'Found end at index {j}')
                break
        
        # Now we need to replace these paragraphs
        # Delete the old ones and insert the new ones
        new_lines = new_passage.split('\n')
        
        # Clear and set the first paragraph
        doc.paragraphs[start_idx].clear()
        doc.paragraphs[start_idx].text = new_lines[0]
        
        # For remaining lines in the new passage, we need to handle this carefully
        # Since we can't easily insert paragraphs, we'll modify the first paragraph
        # to contain all the text
        first_para = doc.paragraphs[start_idx]
        first_para.clear()
        full_text = '\n'.join(new_lines)
        first_para.add_run(full_text)
        
        print('Paragraph updated')
        break

if found:
    # Save the document
    doc.save(docx_path)
    print('✓ DOCX file updated successfully')
else:
    print('Target paragraph not found - checking what we have...')
    for i, para in enumerate(doc.paragraphs):
        if 'central grid flickered' in para.text:
            print(f'Paragraph {i}: {para.text[:100]}...')
