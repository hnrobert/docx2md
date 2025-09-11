"""
Document processing module for handling main document conversion.
"""

from typing import Any, List

from .paragraph_processor import ParagraphProcessor
from .table_processor import TableProcessor

try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)


class DocumentProcessor:
    """Handles main document processing and coordination"""

    def __init__(self, image_extractor, output_lines: List[str]):
        self.output_lines = output_lines
        self.paragraph_processor = ParagraphProcessor(
            image_extractor, output_lines)
        self.table_processor = TableProcessor(output_lines)

    def convert_document(self, doc: Any) -> None:
        """Convert main document content"""
        # First check if there are Title style paragraphs, if so use as main title
        title_found = self._check_for_title_style(doc)

        # Set heading offset: if Title style exists, all headings are adjusted down one level
        heading_offset = 1 if title_found else 0
        self.paragraph_processor.set_heading_offset(heading_offset)

        # Process all document elements
        first_heading_found = False
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                paragraph = Paragraph(element, doc)
                style_name = paragraph.style.name.lower(
                ) if paragraph.style and paragraph.style.name else ''

                # Check Title style
                if 'title' in style_name and paragraph.text.strip():
                    self.output_lines.append(f"# {paragraph.text.strip()}")
                    self.output_lines.append('')
                    continue

                # If no Title, first Heading 1 becomes main title
                if not title_found and not first_heading_found and 'heading 1' in style_name and paragraph.text.strip():
                    self.output_lines.append(f"# {paragraph.text.strip()}")
                    self.output_lines.append('')
                    first_heading_found = True
                    continue

                self.paragraph_processor.convert_paragraph(paragraph)

            elif element.tag.endswith('tbl'):  # Table
                table = Table(element, doc)
                self.table_processor.convert_table(table)

    def _check_for_title_style(self, doc: Any) -> bool:
        """Check if document contains Title style paragraphs"""
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                paragraph = Paragraph(element, doc)
                style_name = paragraph.style.name.lower(
                ) if paragraph.style and paragraph.style.name else ''

                if 'title' in style_name and paragraph.text.strip():
                    return True
        return False
