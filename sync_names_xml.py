import zipfile
import shutil
from lxml import etree

docx_path = r'C:\Users\Prisha\OneDrive\Documents\Miner Town\Book 2\manuscript\chapter-01-assigned-to-witness.docx'

# Create a backup
shutil.copy(docx_path, docx_path + '.bak')

# Open the DOCX (which is a ZIP file)
with zipfile.ZipFile(docx_path, 'r') as zip_ref:
    # Extract document.xml
    doc_xml = zip_ref.read('word/document.xml')

# Parse the XML
root = etree.fromstring(doc_xml)

# Find and replace text
old_text = 'The Initiation queue populated on screen—three entries materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity\'s inventory.'
new_text = 'The Initiation queue populated on screen—the names materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity\'s inventory.'

# Convert to byte string
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# Get all text elements
for t in root.findall('.//w:t', ns):
    if t.text and old_text in t.text:
        t.text = t.text.replace(old_text, new_text)
        print(f'Found and replaced in text element')

# Write back
with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as docx:
    # Re-add document.xml
    docx.writestr('word/document.xml', etree.tostring(root, xml_declaration=True, encoding='UTF-8'))
    
    # Re-add other files from backup
    with zipfile.ZipFile(docx_path + '.bak', 'r') as old_zip:
        for item in old_zip.filelist:
            if item.filename != 'word/document.xml':
                docx.writestr(item, old_zip.read(item.filename))

print('DOCX updated successfully')
