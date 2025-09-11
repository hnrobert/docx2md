"""
Table processing module for converting Word tables to Markdown.
"""

from typing import List

try:
    from docx.table import Table
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)


class TableProcessor:
    """Handles table processing and conversion"""

    def __init__(self, output_lines: List[str]):
        self.output_lines = output_lines

    def convert_table(self, table: Table) -> None:
        """Convert table to Markdown format"""
        self.output_lines.append('')  # Blank line before table

        # Convert table rows
        for i, row in enumerate(table.rows):
            cells = [cell.text.strip().replace('\n', ' ')
                     for cell in row.cells]

            # Table row
            self.output_lines.append('| ' + ' | '.join(cells) + ' |')

            # Add header separator (after first row)
            if i == 0:
                separator = '|' + ''.join([' --- |' for _ in cells])
                self.output_lines.append(separator)

        self.output_lines.append('')  # Blank line after table
