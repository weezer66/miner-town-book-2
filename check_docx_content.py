from docx import Document
from pathlib import Path

docx_path = r'C:\Users\Prisha\OneDrive\Documents\Miner Town\Book 2\manuscript\chapter-01-assigned-to-witness.docx'
markdown_path = Path('manuscript/chapter-01-assigned-to-witness.md')
start_marker = 'Then the central grid flickered.'
end_marker = '"Five feet eight inches. One hundred sixty pounds."'

doc = Document(docx_path)
document_paragraphs = [paragraph.text for paragraph in doc.paragraphs]
document_text = '\n'.join(document_paragraphs)
markdown_text = markdown_path.read_text(encoding='utf-8')


def source_paragraphs(text):
    return [paragraph.replace('**', '') for paragraph in text.strip().split('\n\n')]


markdown_start = markdown_text.index(start_marker)
markdown_end = markdown_text.index(end_marker, markdown_start)
expected_paragraphs = source_paragraphs(markdown_text[markdown_start:markdown_end])

document_start = document_paragraphs.index(start_marker)
document_end = document_paragraphs.index(end_marker, document_start)
actual_paragraphs = [paragraph for paragraph in document_paragraphs[document_start:document_end] if paragraph]

if expected_paragraphs != actual_paragraphs:
    raise SystemExit('Verification failed: DOCX assignment passage differs from Markdown')

if document_text.count(start_marker) != 1:
    raise SystemExit('Verification failed: duplicate assignment passage')

print(f'Verification passed: {len(expected_paragraphs)} assignment paragraphs match Markdown')
