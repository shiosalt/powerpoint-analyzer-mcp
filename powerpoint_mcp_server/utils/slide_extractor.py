"""Extended ZIP extractor for PowerPoint slides with relationship handling."""

import logging
from typing import Dict, Any, Optional
import xml.etree.ElementTree as ET
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)

class SlideExtractor(ZipExtractor):
    """Extended ZIP extractor with PowerPoint relationship handling."""

    def __init__(self, file_path: str):
        """Initialize the slide extractor."""
        super().__init__(file_path)
        self._slide_mappings = None
        self._section_mappings = None

    def get_slide_mappings(self) -> Dict[str, Dict[str, str]]:
        """Get or create slide relationship mappings."""
        if self._slide_mappings is None:
            self._slide_mappings = self._create_slide_mappings()
        return self._slide_mappings

    def get_section_mappings(self) -> list:
        """Get or create section mappings."""
        if self._section_mappings is None:
            self._section_mappings = self._create_section_mappings()
        return self._section_mappings

    def _create_slide_mappings(self) -> Dict[str, Dict[str, str]]:
        """Create mapping of slide relationships."""
        try:
            slide_mapping = {}

            # Read presentation relationships
            rels_content = self.read_xml_content('ppt/_rels/presentation.xml.rels')
            if not rels_content:
                logger.warning("No presentation relationships file found")
                return {}

            # Parse relationships XML
            rels_root = ET.fromstring(rels_content)

            # Find all slide relationships
            rel_namespace = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            slide_rels = rels_root.findall('.//r:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"]', rel_namespace)

            # Create mapping of Target -> r_id
            target_to_rid = {
                rel.get('Target').replace('slides/', ''): rel.get('Id')
                for rel in slide_rels
                if rel.get('Target') and rel.get('Id')
            }

            # Read presentation.xml to get slide IDs and their order
            pres_content = self.read_xml_content('ppt/presentation.xml')
            if not pres_content:
                logger.warning("No presentation.xml file found")
                return {}

            pres_root = ET.fromstring(pres_content)

            # Find all slide entries
            p_namespace = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
            slide_list = pres_root.find('.//p:sldIdLst', p_namespace)

            if slide_list is not None:
                for slide_number, slide_entry in enumerate(slide_list.findall('p:sldId', p_namespace), 1):
                    slide_id = slide_entry.get('id')
                    r_id = slide_entry.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

                    if slide_id and r_id:
                        target_file = f'slide{slide_id}.xml'
                        slide_mapping[slide_id] = {
                            'r_id': r_id,
                            'id': slide_id,
                            'slide_number': str(slide_number)
                        }
                        logger.debug(f"Mapped slide {slide_number}: id={slide_id}, r_id={r_id}")

            return slide_mapping

        except Exception as e:
            logger.error(f"Failed to create slide mappings: {e}")
            return {}

    def _create_section_mappings(self) -> list:
        """Create section mappings with slide information."""
        try:
            sections = []
            slide_mappings = self.get_slide_mappings()

            # Read presentation.xml
            pres_content = self.read_xml_content('ppt/presentation.xml')
            if not pres_content:
                logger.warning("No presentation.xml file found")
                return []

            pres_root = ET.fromstring(pres_content)

            # Find section list
            p_namespace = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
            section_list = pres_root.find('.//p:sectionLst', p_namespace)

            if section_list is not None:
                for section in section_list.findall('p:section', p_namespace):
                    section_name = section.get('name', 'Unnamed Section')
                    section_id = section.get('id', '')

                    # Get slides in this section
                    slide_list = section.find('p:sldIdLst', p_namespace)
                    slide_info = []

                    if slide_list is not None:
                        for slide_ref in slide_list.findall('p:sldId', p_namespace):
                            slide_id = slide_ref.get('id')
                            if slide_id and slide_id in slide_mappings:
                                info = {
                                    'id': slide_id,
                                    'slide_number': slide_mappings[slide_id]['slide_number']
                                }
                                slide_info.append(info)

                    sections.append({
                        'name': section_name,
                        'id': section_id,
                        'slide_ids': slide_info,
                        'slide_count': len(slide_info)
                    })

                    logger.debug(f"Section '{section_name}' has {len(slide_info)} slides")

            return sections

        except Exception as e:
            logger.error(f"Failed to create section mappings: {e}")
            return []
