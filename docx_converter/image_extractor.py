"""
Image extraction module for DOCX files.
"""

import logging
import os
import shutil
import xml.etree.ElementTree as ET
import zipfile
from typing import Dict, Optional

logger = logging.getLogger(__name__)


class ImageExtractor:
    """Handles image extraction from DOCX files"""

    def __init__(self, assets_dir: str):
        self.assets_dir = assets_dir
        self.image_counter = 0
        self.image_map: Dict[str, str] = {}

    def extract_images(self, docx_path: str) -> None:
        """
        Extract images from DOCX file and establish mapping relationship

        Args:
            docx_path: Path to the DOCX file
        """
        if not self.assets_dir:
            return

        try:
            # Reset image counter and mapping
            self.image_counter = 0
            self.image_map = {}

            # DOCX file is actually a ZIP file
            with zipfile.ZipFile(docx_path, 'r') as docx_zip:
                # Read relationship file to get image relationship mapping
                try:
                    rels_content = docx_zip.read(
                        'word/_rels/document.xml.rels').decode('utf-8')
                    rels_root = ET.fromstring(rels_content)

                    # Establish relationship ID to image file mapping
                    self._extract_images_with_relationships(
                        docx_zip, rels_root)

                except Exception as e:
                    logger.warning(
                        f"Unable to parse image relationships, using fallback method: {e}")
                    # Fallback method: directly extract all images from media folder
                    self._extract_images_fallback(docx_zip)

        except Exception as e:
            logger.warning(f"Error extracting images: {str(e)}")

    def _extract_images_with_relationships(self, docx_zip: zipfile.ZipFile, rels_root: ET.Element) -> None:
        """Extract images using relationship mapping"""
        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rel_type = rel.get('Type', '')
            if 'image' in rel_type.lower():
                rel_id = rel.get('Id')
                target = rel.get('Target')
                if target and target.startswith('media/'):
                    full_path = f"word/{target}"
                    if full_path in [f.filename for f in docx_zip.filelist]:
                        # Extract image
                        self.image_counter += 1
                        file_ext = os.path.splitext(target)[1].lower()
                        new_filename = f"image_{self.image_counter:03d}{file_ext}"
                        output_path = os.path.join(
                            self.assets_dir, new_filename)

                        with docx_zip.open(full_path) as source:
                            with open(output_path, 'wb') as target_file:
                                shutil.copyfileobj(source, target_file)

                        # Establish mapping relationship
                        if rel_id:
                            self.image_map[rel_id] = new_filename
                            logger.info(
                                f"Extracted image: {new_filename} (ID: {rel_id})")

    def _extract_images_fallback(self, docx_zip: zipfile.ZipFile) -> None:
        """Fallback method to extract images"""
        for file_info in docx_zip.filelist:
            if file_info.filename.startswith('word/media/'):
                file_ext = os.path.splitext(file_info.filename)[1].lower()
                if file_ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg']:
                    self.image_counter += 1
                    new_filename = f"image_{self.image_counter:03d}{file_ext}"
                    output_path = os.path.join(self.assets_dir, new_filename)

                    with docx_zip.open(file_info.filename) as source:
                        with open(output_path, 'wb') as target:
                            shutil.copyfileobj(source, target)

                    logger.info(f"Extracted image: {new_filename}")

    def get_image_reference(self, rel_id: Optional[str] = None) -> str:
        """
        Get image reference for Markdown

        Args:
            rel_id: Relationship ID of the image

        Returns:
            Markdown image reference string
        """
        if rel_id and rel_id in self.image_map:
            image_filename = self.image_map[rel_id]
            return f"![Image](./assets/{image_filename})"
        elif self.image_counter > 0:
            # Use generic image reference
            image_filename = f"image_001.png"
            return f"![Image](./assets/{image_filename})"
        else:
            return ""

    def has_images(self) -> bool:
        """Check if any images were extracted"""
        return self.image_counter > 0
