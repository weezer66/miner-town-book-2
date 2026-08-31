from html.parser import HTMLParser
from pathlib import Path
from textwrap import fill
from zipfile import ZipFile


class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        self.parts.append(data)


book_path = Path(r"C:\Users\Prisha\OneDrive\Documents\Miner Town\Amazon KDP\Miner_Town_Awakening_KDP.epub")
output_parts = []

with ZipFile(book_path) as archive:
    chapter_paths = [
        name
        for name in archive.namelist()
        if name.lower().endswith((".xhtml", ".html", ".htm"))
    ]
    for chapter_path in chapter_paths:
        if not chapter_path.startswith(("EPUB/text/ch004", "EPUB/text/ch005")):
            continue
        parser = TextExtractor()
        parser.feed(archive.read(chapter_path).decode("utf-8"))
        text = " ".join(part.strip() for part in parser.parts if part.strip())
        output_parts.append(f"\n--- {chapter_path} ---\n{fill(text, width=100)}\n")

    Path("_book1_chapters_3_4.txt").write_text("".join(output_parts), encoding="utf-8")
