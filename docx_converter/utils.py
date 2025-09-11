"""
Utility functions for DOCX to Markdown conversion.
"""

import re
from typing import List


def clean_markdown_content(output_lines: List[str]) -> str:
    """
    Clean and format Markdown content

    Args:
        output_lines: List of output lines

    Returns:
        Cleaned Markdown content string
    """
    # Generate Markdown content
    markdown_content = '\n'.join(output_lines)

    # Clean up extra blank lines - merge multiple consecutive blank lines into single blank line
    markdown_content = re.sub(r'\n{3,}', '\n\n', markdown_content)

    # Remove blank lines at beginning and end
    markdown_content = markdown_content.strip()

    # Add extra blank line at the end
    markdown_content += '\n'

    return markdown_content


def extract_heading_level(style_name: str) -> int:
    """Extract heading level from style name"""
    match = re.search(r'heading\s*(\d+)', style_name)
    if match:
        # Markdown supports maximum 6 heading levels
        return min(int(match.group(1)), 6)
    return 1


def merge_adjacent_tags(text: str) -> str:
    """Merge adjacent HTML tags of the same type"""
    # Merge adjacent underline tags
    text = re.sub(r'</u><u>', '', text)
    return text


def is_list_marker_text(text: str) -> bool:
    """Check if text starts with list markers"""
    list_markers = ['•', '◦', '▪', '▫', '‣', '-', '*', '+']
    return any(text.startswith(marker + ' ') for marker in list_markers)


def is_numbered_list_text(text: str) -> bool:
    """Check if text is a numbered list"""
    return bool(re.match(r'^\d+[\.）]\s+', text))


def remove_list_markers(text: str) -> str:
    """Remove list markers from text"""
    # Remove numbered list markers
    text = re.sub(r'^\d+[\.）]\s+', '', text)

    # Remove unordered list markers
    list_markers = ['•', '◦', '▪', '▫', '‣', '-', '*', '+']
    for marker in list_markers:
        if text.startswith(marker + ' '):
            text = text[len(marker):].strip()
            break

    return text
