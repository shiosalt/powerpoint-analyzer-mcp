"""
Position-aware text formatting extraction for PowerPoint slides.
"""

import re
import logging
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

from .content_extractor import ContentExtractor
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)


@dataclass
class FormattingSegment:
    """Represents a formatted text segment with precise positioning."""
    text: str
    start_position: int
    end_position: int
    formatting_type: str
    element_index: int
    additional_info: Optional[Dict[str, Any]] = None


@dataclass
class SlideFormattingResult:
    """Formatting results for a single slide."""
    slide_number: int
    title: str
    complete_text: str
    formatted_segments: List[FormattingSegment]


@dataclass
class FormattingExtractionResult:
    """Result of formatting extraction with position-aware segments."""
    file_path: str
    formatting_type: str
    summary: Dict[str, int]
    results_by_slide: List[SlideFormattingResult]


class FormattingExtractor:
    """
    Position-aware text formatting extractor for PowerPoint slides.
    """
    
    def __init__(self, content_extractor: Optional[ContentExtractor] = None):
        """Initialize the formatting extractor."""
        self.content_extractor = content_extractor or ContentExtractor()
        self._extraction_cache = {}
    
    def extract_formatting(
        self,
        file_path: str,
        formatting_type: str,
        slide_numbers: Optional[List[int]] = None
    ) -> FormattingExtractionResult:
        """
        Extract text with specific formatting attributes from PowerPoint slides.
        
        Args:
            file_path: Path to the PowerPoint file
            formatting_type: Type of formatting to extract
            slide_numbers: List of slide numbers to analyze (None for all)
            
        Returns:
            FormattingExtractionResult with position-aware segments
        """
        logger.info(f"Extracting {formatting_type} formatting from {file_path}")
        
        # Validate formatting type
        valid_types = ['bold', 'italic', 'underlined', 'highlighted', 'strikethrough', 
                      'hyperlinks', 'font_sizes', 'font_colors']
        if formatting_type not in valid_types:
            raise ValueError(f"Invalid formatting type '{formatting_type}'. Valid types: {valid_types}")
        
        try:
            results_by_slide = []
            
            with ZipExtractor(file_path) as extractor:
                # Get slide XML files sorted numerically
                slide_files = extractor.get_slide_xml_files_sorted()
                
                # Determine which slides to analyze
                if slide_numbers is None:
                    slides_to_analyze = list(range(1, len(slide_files) + 1))
                else:
                    slides_to_analyze = [s for s in slide_numbers if s <= len(slide_files)]
                
                for slide_num in slides_to_analyze:
                    slide_file = slide_files[slide_num - 1]
                    slide_xml = extractor.read_xml_content(slide_file)
                    
                    if slide_xml:
                        slide_result = self._extract_formatting_from_slide(
                            slide_xml, slide_num, formatting_type, extractor
                        )
                        results_by_slide.append(slide_result)
            
            # Create summary
            summary = self._create_extraction_summary(results_by_slide, formatting_type)
            
            result = FormattingExtractionResult(
                file_path=file_path,
                formatting_type=formatting_type,
                summary=summary,
                results_by_slide=results_by_slide
            )
            
            logger.info(f"Extracted {summary['total_formatted_segments']} {formatting_type} segments from {len(results_by_slide)} slides")
            return result
            
        except Exception as e:
            logger.error(f"Error extracting {formatting_type} formatting: {e}")
            raise
    
    def _extract_formatting_from_slide(
        self,
        slide_xml: str,
        slide_number: int,
        formatting_type: str,
        extractor: ZipExtractor
    ) -> SlideFormattingResult:
        """Extract formatting from a single slide."""
        try:
            root = self.content_extractor.xml_parser.parse_xml_string(slide_xml)
            if root is None:
                return SlideFormattingResult(
                    slide_number=slide_number,
                    title="",
                    complete_text="",
                    formatted_segments=[]
                )
            
            # Extract slide title
            title = self._extract_slide_title(root)
            
            # Extract complete text and build formatting segments
            complete_text, formatted_segments = self._extract_text_with_formatting(
                root, formatting_type, extractor
            )
            
            return SlideFormattingResult(
                slide_number=slide_number,
                title=title,
                complete_text=complete_text,
                formatted_segments=formatted_segments
            )
            
        except Exception as e:
            logger.warning(f"Failed to extract formatting from slide {slide_number}: {e}")
            return SlideFormattingResult(
                slide_number=slide_number,
                title="",
                complete_text="",
                formatted_segments=[]
            )
    
    def _extract_slide_title(self, root) -> str:
        """Extract slide title."""
        try:
            # Find title placeholders
            shapes = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:sp'
            )
            
            for shape in shapes:
                nv_sp_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                    shape, './/p:nvSpPr'
                )
                if nv_sp_pr is not None:
                    ph = self.content_extractor.xml_parser.find_element_with_namespace(
                        nv_sp_pr, './/p:ph'
                    )
                    if ph is not None and ph.get('type') == 'title':
                        return self.content_extractor._extract_shape_text_content(shape)
            
            return ""
            
        except Exception as e:
            logger.warning(f"Failed to extract slide title: {e}")
            return ""
    
    def _extract_text_with_formatting(
        self,
        root,
        formatting_type: str,
        extractor: ZipExtractor
    ) -> Tuple[str, List[FormattingSegment]]:
        """Extract complete text and identify formatted segments."""
        try:
            complete_text_parts = []
            formatted_segments = []
            current_position = 0
            element_index = 0
            
            # Find all text-containing shapes
            shapes = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:sp'
            )
            
            for shape in shapes:
                tx_body = self.content_extractor.xml_parser.find_element_with_namespace(
                    shape, './/p:txBody'
                )
                
                if tx_body is not None:
                    shape_text, shape_segments = self._extract_shape_text_with_formatting(
                        tx_body, formatting_type, current_position, element_index
                    )
                    
                    if shape_text:
                        complete_text_parts.append(shape_text)
                        formatted_segments.extend(shape_segments)
                        current_position += len(shape_text)
                        element_index += 1
            
            # Also check tables
            graphic_frames = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:graphicFrame'
            )
            
            for frame in graphic_frames:
                table_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    frame, './/a:tbl'
                )
                
                if table_elem is not None:
                    table_text, table_segments = self._extract_table_text_with_formatting(
                        table_elem, formatting_type, current_position, element_index
                    )
                    
                    if table_text:
                        complete_text_parts.append(table_text)
                        formatted_segments.extend(table_segments)
                        current_position += len(table_text)
                        element_index += 1
            
            complete_text = " ".join(complete_text_parts)
            
            # Handle hyperlinks specially as they need relationship resolution
            if formatting_type == 'hyperlinks':
                hyperlink_segments = self._extract_hyperlink_segments(
                    root, complete_text, extractor
                )
                formatted_segments.extend(hyperlink_segments)
            
            return complete_text, formatted_segments
            
        except Exception as e:
            logger.warning(f"Failed to extract text with formatting: {e}")
            return "", []
    
    def _extract_shape_text_with_formatting(
        self,
        tx_body,
        formatting_type: str,
        base_position: int,
        element_index: int
    ) -> Tuple[str, List[FormattingSegment]]:
        """Extract text and formatting from a text body."""
        try:
            text_parts = []
            formatted_segments = []
            current_pos = 0
            
            # Find all paragraphs
            paragraphs = self.content_extractor.xml_parser.find_elements_with_namespace(
                tx_body, './/a:p'
            )
            
            for para in paragraphs:
                # Find all runs in the paragraph
                runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                    para, './/a:r'
                )
                
                for run in runs:
                    # Extract text from run
                    text_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        run, './/a:t'
                    )
                    
                    if text_elem is not None and text_elem.text:
                        run_text = text_elem.text
                        text_parts.append(run_text)
                        
                        # Check if this run has the requested formatting
                        if self._run_has_formatting(run, formatting_type):
                            segment = FormattingSegment(
                                text=run_text,
                                start_position=base_position + current_pos,
                                end_position=base_position + current_pos + len(run_text),
                                formatting_type=formatting_type,
                                element_index=element_index,
                                additional_info=self._get_additional_formatting_info(run, formatting_type)
                            )
                            formatted_segments.append(segment)
                        
                        current_pos += len(run_text)
                
                # Add paragraph break
                if paragraphs.index(para) < len(paragraphs) - 1:
                    text_parts.append(" ")
                    current_pos += 1
            
            shape_text = "".join(text_parts)
            return shape_text, formatted_segments
            
        except Exception as e:
            logger.warning(f"Failed to extract shape text with formatting: {e}")
            return "", []
    
    def _extract_table_text_with_formatting(
        self,
        table_elem,
        formatting_type: str,
        base_position: int,
        element_index: int
    ) -> Tuple[str, List[FormattingSegment]]:
        """Extract text and formatting from a table."""
        try:
            text_parts = []
            formatted_segments = []
            current_pos = 0
            
            # Find all table rows
            rows = self.content_extractor.xml_parser.find_elements_with_namespace(
                table_elem, './/a:tr'
            )
            
            for row in rows:
                # Find all cells in the row
                cells = self.content_extractor.xml_parser.find_elements_with_namespace(
                    row, './/a:tc'
                )
                
                for cell in cells:
                    # Extract text from cell
                    cell_text = self.content_extractor._extract_cell_text_content(cell)
                    if cell_text:
                        text_parts.append(cell_text)
                        
                        # Check cell formatting
                        cell_segments = self._extract_cell_formatting_segments(
                            cell, formatting_type, base_position + current_pos, element_index
                        )
                        formatted_segments.extend(cell_segments)
                        
                        current_pos += len(cell_text)
                        
                        # Add cell separator
                        text_parts.append(" ")
                        current_pos += 1
                
                # Add row separator
                text_parts.append(" ")
                current_pos += 1
            
            table_text = "".join(text_parts)
            return table_text, formatted_segments
            
        except Exception as e:
            logger.warning(f"Failed to extract table text with formatting: {e}")
            return "", []
    
    def _extract_cell_formatting_segments(
        self,
        cell,
        formatting_type: str,
        base_position: int,
        element_index: int
    ) -> List[FormattingSegment]:
        """Extract formatting segments from a table cell."""
        try:
            segments = []
            current_pos = 0
            
            # Find text body in cell
            tx_body = self.content_extractor.xml_parser.find_element_with_namespace(
                cell, './/a:txBody'
            )
            
            if tx_body is not None:
                # Find all runs in the cell
                runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                    tx_body, './/a:r'
                )
                
                for run in runs:
                    text_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        run, './/a:t'
                    )
                    
                    if text_elem is not None and text_elem.text:
                        run_text = text_elem.text
                        
                        if self._run_has_formatting(run, formatting_type):
                            segment = FormattingSegment(
                                text=run_text,
                                start_position=base_position + current_pos,
                                end_position=base_position + current_pos + len(run_text),
                                formatting_type=formatting_type,
                                element_index=element_index,
                                additional_info=self._get_additional_formatting_info(run, formatting_type)
                            )
                            segments.append(segment)
                        
                        current_pos += len(run_text)
            
            return segments
            
        except Exception as e:
            logger.warning(f"Failed to extract cell formatting segments: {e}")
            return []
    
    def _run_has_formatting(self, run, formatting_type: str) -> bool:
        """Check if a text run has the specified formatting."""
        try:
            r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                run, './/a:rPr'
            )
            
            if r_pr is None:
                return False
            
            if formatting_type == 'bold':
                bold_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:b'
                )
                if bold_elem is not None:
                    return bold_elem.get('val', '1') != '0'
            
            elif formatting_type == 'italic':
                italic_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:i'
                )
                if italic_elem is not None:
                    return italic_elem.get('val', '1') != '0'
            
            elif formatting_type == 'underlined':
                underline_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:u'
                )
                if underline_elem is not None:
                    return underline_elem.get('val', 'sng') != 'none'
            
            elif formatting_type == 'highlighted':
                highlight_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:highlight'
                )
                return highlight_elem is not None
            
            elif formatting_type == 'strikethrough':
                strike_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:strike'
                )
                if strike_elem is not None:
                    return strike_elem.get('val', 'sngStrike') != 'noStrike'
            
            elif formatting_type == 'font_sizes':
                font_size_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:sz'
                )
                return font_size_elem is not None
            
            elif formatting_type == 'font_colors':
                solid_fill = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:solidFill'
                )
                return solid_fill is not None
            
            return False
            
        except Exception as e:
            logger.warning(f"Failed to check run formatting: {e}")
            return False
    
    def _get_additional_formatting_info(self, run, formatting_type: str) -> Optional[Dict[str, Any]]:
        """Get additional formatting information for a run."""
        try:
            r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                run, './/a:rPr'
            )
            
            if r_pr is None:
                return None
            
            info = {}
            
            if formatting_type == 'font_sizes':
                font_size_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:sz'
                )
                if font_size_elem is not None:
                    sz = font_size_elem.get('val')
                    if sz:
                        try:
                            info['font_size'] = float(sz) / 100.0
                        except (ValueError, TypeError):
                            pass
            
            elif formatting_type == 'font_colors':
                solid_fill = self.content_extractor.xml_parser.find_element_with_namespace(
                    r_pr, './/a:solidFill'
                )
                if solid_fill is not None:
                    color = self._extract_color_from_fill(solid_fill)
                    if color:
                        info['color'] = color
            
            return info if info else None
            
        except Exception as e:
            logger.warning(f"Failed to get additional formatting info: {e}")
            return None
    
    def _extract_color_from_fill(self, solid_fill) -> Optional[str]:
        """Extract color value from a solid fill element."""
        try:
            # Look for RGB color
            srgb_clr = self.content_extractor.xml_parser.find_element_with_namespace(
                solid_fill, './/a:srgbClr'
            )
            if srgb_clr is not None:
                color_val = srgb_clr.get('val')
                if color_val:
                    return f"#{color_val}"
            
            # Look for scheme color
            scheme_clr = self.content_extractor.xml_parser.find_element_with_namespace(
                solid_fill, './/a:schemeClr'
            )
            if scheme_clr is not None:
                color_val = scheme_clr.get('val')
                if color_val:
                    return color_val
            
            return None
            
        except Exception as e:
            logger.warning(f"Failed to extract color from fill: {e}")
            return None
    
    def _extract_hyperlink_segments(
        self,
        root,
        complete_text: str,
        extractor: ZipExtractor
    ) -> List[FormattingSegment]:
        """Extract hyperlink segments with relationship resolution."""
        try:
            segments = []
            
            # Find all hyperlinks
            hyperlinks = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/a:hlinkClick'
            )
            
            for hl in hyperlinks:
                # Get relationship ID
                r_id = hl.get('id') or hl.get('r:id') or hl.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                
                if r_id:
                    # Find the text associated with this hyperlink
                    # This is a simplified approach - in reality, we'd need to track
                    # the exact text runs that contain the hyperlink
                    hyperlink_text = self._find_hyperlink_text(hl, root)
                    
                    if hyperlink_text:
                        # Find position in complete text
                        start_pos = complete_text.find(hyperlink_text)
                        if start_pos >= 0:
                            segment = FormattingSegment(
                                text=hyperlink_text,
                                start_position=start_pos,
                                end_position=start_pos + len(hyperlink_text),
                                formatting_type='hyperlinks',
                                element_index=0,
                                additional_info={'relationship_id': r_id}
                            )
                            segments.append(segment)
            
            return segments
            
        except Exception as e:
            logger.warning(f"Failed to extract hyperlink segments: {e}")
            return []
    
    def _find_hyperlink_text(self, hyperlink_elem, root) -> str:
        """Find the text associated with a hyperlink element."""
        try:
            # This is a simplified implementation
            # In practice, we'd need to traverse the XML structure more carefully
            # to find the exact text runs that contain the hyperlink
            
            # For now, return a placeholder
            return "hyperlink_text"
            
        except Exception as e:
            logger.warning(f"Failed to find hyperlink text: {e}")
            return ""
    
    def _create_extraction_summary(
        self,
        results_by_slide: List[SlideFormattingResult],
        formatting_type: str
    ) -> Dict[str, int]:
        """Create summary statistics for the extraction."""
        try:
            total_slides_analyzed = len(results_by_slide)
            slides_with_formatting = 0
            total_formatted_segments = 0
            
            for slide_result in results_by_slide:
                if slide_result.formatted_segments:
                    slides_with_formatting += 1
                    total_formatted_segments += len(slide_result.formatted_segments)
            
            return {
                'total_slides_analyzed': total_slides_analyzed,
                'slides_with_formatting': slides_with_formatting,
                'total_formatted_segments': total_formatted_segments
            }
            
        except Exception as e:
            logger.warning(f"Failed to create extraction summary: {e}")
            return {
                'total_slides_analyzed': 0,
                'slides_with_formatting': 0,
                'total_formatted_segments': 0
            }