"""
Image processing module for handling images in paragraphs.
"""

import logging

try:
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Error: Missing required library. Please run: pip install python-docx")
    import sys
    sys.exit(1)

logger = logging.getLogger(__name__)


class ImageProcessor:
    """Handles image processing in paragraphs"""

    def __init__(self, image_extractor):
        self.image_extractor = image_extractor

    def process_paragraph_images(self, paragraph: Paragraph) -> str:
        """
        Process images in paragraph

        Args:
            paragraph: Word paragraph object

        Returns:
            Markdown image references as string
        """
        images_found = []

        # Check if paragraph contains image elements
        para_element = paragraph._element

        # Method 1: Find w:drawing elements (new image format)
        drawings = para_element.xpath('.//w:drawing')
        logger.debug(f"Found {len(drawings)} drawing elements in paragraph")

        for drawing in drawings:
            # Find image relationship ID - using simplified method
            blip_elements = []
            # Iterate through all elements in drawing to find blip elements
            for elem in drawing.iter():
                if elem.tag.endswith('}blip'):
                    blip_elements.append(elem)

            for blip in blip_elements:
                rel_id = blip.get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                logger.debug(f"Found image relationship ID: {rel_id}")

                image_ref = self.image_extractor.get_image_reference(rel_id)
                if image_ref:
                    images_found.append(image_ref)
                    logger.info(f"Inserted image link for ID: {rel_id}")

            # If no blip elements found but have drawing, indicates there are images
            if not blip_elements and self.image_extractor.has_images():
                image_ref = self.image_extractor.get_image_reference()
                if image_ref:
                    images_found.append(image_ref)
                    logger.info("Using fallback image link")

        # Method 2: Find w:pict elements (old image format)
        picts = para_element.xpath('.//w:pict')
        logger.debug(f"Found {len(picts)} pict elements in paragraph")

        for pict in picts:
            # Create reference for old image format
            if self.image_extractor.has_images():
                image_ref = self.image_extractor.get_image_reference()
                if image_ref:
                    images_found.append(image_ref)
                    logger.info("Inserted old image link")

        # Method 3: Check images in runs
        for run in paragraph.runs:
            run_drawings = run._element.xpath('.//w:drawing')
            run_picts = run._element.xpath('.//w:pict')

            if run_drawings or run_picts:
                logger.debug(
                    f"Found image elements in run: drawings={len(run_drawings)}, picts={len(run_picts)}")

                # If no images found earlier but there are image elements here
                if not images_found and self.image_extractor.has_images():
                    image_ref = self.image_extractor.get_image_reference()
                    if image_ref:
                        images_found.append(image_ref)
                        logger.info("Inserted image link in run")

        if images_found:
            logger.info(f"Total {len(images_found)} images found in paragraph")

        return '\n'.join(images_found) if images_found else ""
