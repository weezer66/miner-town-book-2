from docx import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

docx_path = r'C:\Users\Prisha\OneDrive\Documents\Miner Town\Book 2\manuscript\chapter-01-assigned-to-witness.docx'

doc = Document(docx_path)

for paragraph in doc.paragraphs:
    if paragraph.text.startswith('The Initiation queue populated on screen—sixteen names materialized'):
        paragraph.text = "The Initiation queue populated on screen—sixteen names materialized like component numbers in a system ledger, stripped of color or ceremony, rendered as mere data points in Trinity's inventory. Sixteen children. Four monitors. Four ganas. On paper, the mathematics promised a clean division. Predictable."
    elif paragraph.text.startswith('Elara leaned toward the screen as the system began its assignment sequence.'):
        paragraph.text = "Elara leaned toward the screen as the system began its assignment sequence. Thirteen names streamed toward the EVENT, THRESHOLD, CORRECTION, and VARIANCE consoles, each connected to its designated gana with the mechanical precision Trinity preferred. At first, the grid moved through the queue with invisible efficiency."
    elif paragraph.text.startswith('Three names remained on the primary display, orphaned.'):
        paragraph.text = 'Three names remained on the primary display, never routed to a console. And beneath each name, where the system normally displayed the responsible gana, a fifth designation appeared instead.'
    elif paragraph.text.startswith("This was probably the first time Trinity's surveillance system had encountered an anomaly"):
        paragraph.text = "This was probably the first time Trinity's surveillance system had encountered an anomaly of this kind. The four gana representatives searched the online archives and the printed manuals beside their consoles, but neither offered a reason for the designation or an instruction for what came next."

doc.save(docx_path)
print('DOCX file updated successfully')
