"""
Text formatting module for converting Word formatting to Markdown.
"""

from typing import Optional

from .utils import merge_adjacent_tags

try:
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)


class TextFormatter:
    """Handles text formatting conversion from Word to Markdown"""

    def convert_paragraph_formatting(self, paragraph: Paragraph, custom_text: Optional[str] = None) -> str:
        """
        Convert paragraph formatting (bold, italic, links, etc.)

        Args:
            paragraph: Word paragraph object
            custom_text: Custom text to use instead of paragraph runs

        Returns:
            Formatted Markdown text
        """
        if custom_text:
            # If custom text is provided, use simplified processing
            return custom_text

        result = []
        for run in paragraph.runs:
            text = run.text
            if not text:
                continue

            # Apply formatting
            if run.bold:
                text = f"**{text}**"
            if run.italic:
                text = f"*{text}*"
            if run.underline:
                text = f"<u>{text}</u>"

            result.append(text)

        # Merge adjacent same HTML tags
        final_result = ''.join(result)

        # Merge adjacent tags of same type
        final_result = merge_adjacent_tags(final_result)

        return final_result
