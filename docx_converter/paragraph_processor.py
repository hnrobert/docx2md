"""
Paragraph processing module for DOCX to Markdown conversion.
"""

import logging
from typing import List, Optional

from .formatting import TextFormatter
from .image_processor import ImageProcessor
from .list_processor import ListProcessor
from .utils import extract_heading_level

try:
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)

logger = logging.getLogger(__name__)


class ParagraphProcessor:
    """Handles paragraph processing and conversion"""

    def __init__(self, image_extractor, output_lines: List[str]):
        self.output_lines = output_lines
        self.text_formatter = TextFormatter()
        self.image_processor = ImageProcessor(image_extractor)
        self.list_processor = ListProcessor(output_lines, self.text_formatter)
        self.heading_offset = 0

    def set_heading_offset(self, offset: int):
        """Set heading level offset"""
        self.heading_offset = offset

    def convert_paragraph(self, paragraph: Paragraph) -> None:
        """Convert paragraph to Markdown"""
        # Get paragraph text
        text = paragraph.text.strip()

        # First check if paragraph contains images (regardless of text content)
        images_text = self.image_processor.process_paragraph_images(paragraph)

        # If paragraph is mainly images (no text or very little text)
        if images_text and (not text or len(text) < 3):
            self.output_lines.append(images_text)
            self.output_lines.append('')
            return

        # Skip empty paragraphs but keep one blank line for separation
        if not text and not images_text:
            # If previous line is not empty, add blank line
            if self.output_lines and self.output_lines[-1] != '':
                self.output_lines.append('')
            return

        # Check paragraph style
        style_name = paragraph.style.name.lower(
        ) if paragraph.style and paragraph.style.name else ''

        # Skip Title style, already handled in document processor
        if 'title' in style_name:
            return

        # Check if it's a list item
        is_list = self.list_processor.is_list_paragraph(paragraph)

        # If previously in list but current is not list item, list ends
        if self.list_processor.in_list and not is_list:
            # Add blank line to separate list and subsequent content
            self.output_lines.append('')
            self.list_processor.end_list()

        # Handle headings (adjust level based on Title style presence)
        if 'heading' in style_name:
            self._convert_heading(paragraph, text, style_name)
            return

        # Handle lists
        if is_list:
            self.list_processor.convert_list_item(paragraph)
            return

        # Handle regular paragraphs
        # If paragraph contains images, insert images first
        if images_text:
            self.output_lines.append(images_text)
            self.output_lines.append('')

        # Handle text content
        if text:  # Only process when paragraph has text
            markdown_text = self.text_formatter.convert_paragraph_formatting(
                paragraph)
            self.output_lines.append(markdown_text)
            self.output_lines.append('')

    def _convert_heading(self, paragraph: Paragraph, text: str, style_name: str) -> None:
        """Convert heading paragraph"""
        level = extract_heading_level(style_name)

        # If Title style exists, all headings are adjusted down one level
        level += self.heading_offset

        # Ensure not exceeding 6 heading levels
        level = min(level, 6)

        self.output_lines.append(f"{'#' * level} {text}")
        self.output_lines.append('')
