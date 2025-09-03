"""
Enhanced formatting extraction system for PowerPoint slides.
Provides position tracking and support for multiple formatting types.
"""

import re
import logging
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass
from enum import Enum

from .content_extractor import ContentExtractor
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)


class FormattingType(Enum):
    """Enumeration of supported formatting types."""
    BOLD = "bold"
    ITALIC = "italic"
    UNDERLINED = "underlined"
    HIGHLIGHTED = "highlighted"
    STRIKETHROUGH = "strikethrough"
    HYPERLINKS = "hyperlinks"
    FONT_SIZES = "font_sizes"
    FONT_COLORS = "font_colors"


@dataclass
class FormattedSegment:
    """Represents a formatted text segment with position."""
    text: str
    start_position: int
    end_position: int
    formatting_type: str
    formatting_details: Dict[str, Any]
    slide_number: int
    element_index: int


@dataclass
class HyperlinkSegment(FormattedSegment):
    """Hyperlink-specific formatted segment."""
    url: str
    display_text: str
    link_type: str  # "external", "internal", "email"


class FormattingExtractor:
    """Enhanced formatting extraction with position tracking."""
    
    def __init__(self, content_extractor: Optional[ContentExtractor] = None):
        """Initialize the formatting extractor."""
        self.content_extractor = content_extractor or ContentExtractor()
        self._cache = {}
    
    def extract_formatting_segments(
        self, 
        text_elements: List[Dict], 
        formatting_type: str,
        slide_number: int
    ) -> List[FormattedSegment]:
        """Extract formatted segments with position information."""
        try:
            if formatting_type not in [ft.value for ft in FormattingType]:
                raise ValueError(f"Invalid formatting_type: {formatting_type}. Valid options: {[ft.value for ft in FormattingType]}")
            
            segments = []
            
            for element_index, text_element in enumerate(text_elements):
                complete_text = text_element.get('content_plain', '')
                if not complete_text:
                    continue
                
                if formatting_type == FormattingType.BOLD.value:
                    segments.extend(self._extract_bold_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
                elif formatting_type == FormattingType.ITALIC.value:
                    segments.extend(self._extract_italic_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
                elif formatting_type == FormattingType.UNDERLINED.value:
                    segments.extend(self._extract_underlined_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
                elif formatting_type == FormattingType.HIGHLIGHTED.value:
                    segments.extend(self._extract_highlighted_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
                elif formatting_type == FormattingType.STRIKETHROUGH.value:
                    segments.extend(self._extract_strikethrough_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
                elif formatting_type == FormattingType.HYPERLINKS.value:
                    segments.extend(self._extract_hyperlink_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
                elif formatting_type == FormattingType.FONT_SIZES.value:
                    segments.extend(self._extract_font_size_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
                elif formatting_type == FormattingType.FONT_COLORS.value:
                    segments.extend(self._extract_font_color_segments(
                        text_element, complete_text, slide_number, element_index
                    ))
            
            return segments
            
        except Exception as e:
            logger.error(f"Error extracting formatting segments: {e}")
            return []
    
    def _extract_bold_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[FormattedSegment]:
        """Extract bold text segments."""
        segments = []
        
        # Check if element has bold formatting
        bold_count = text_element.get('bolded', 0)
        if bold_count > 0:
            # For now, treat the entire text as bold if any bold formatting is detected
            # In a more sophisticated implementation, we would parse the XML runs
            segments.append(FormattedSegment(
                text=complete_text,
                start_position=0,
                end_position=len(complete_text),
                formatting_type=FormattingType.BOLD.value,
                formatting_details={'bold_count': bold_count},
                slide_number=slide_number,
                element_index=element_index
            ))
        
        return segments
    
    def _extract_italic_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[FormattedSegment]:
        """Extract italic text segments."""
        segments = []
        
        # Check if element has italic formatting
        italic_count = text_element.get('italicized', 0)
        if italic_count > 0:
            segments.append(FormattedSegment(
                text=complete_text,
                start_position=0,
                end_position=len(complete_text),
                formatting_type=FormattingType.ITALIC.value,
                formatting_details={'italic_count': italic_count},
                slide_number=slide_number,
                element_index=element_index
            ))
        
        return segments
    
    def _extract_underlined_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[FormattedSegment]:
        """Extract underlined text segments."""
        segments = []
        
        # Check if element has underline formatting
        underlined_count = text_element.get('underlined', 0)
        if underlined_count > 0:
            segments.append(FormattedSegment(
                text=complete_text,
                start_position=0,
                end_position=len(complete_text),
                formatting_type=FormattingType.UNDERLINED.value,
                formatting_details={'underlined_count': underlined_count},
                slide_number=slide_number,
                element_index=element_index
            ))
        
        return segments
    
    def _extract_highlighted_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[FormattedSegment]:
        """Extract highlighted text segments."""
        segments = []
        
        # Check if element has highlight formatting
        highlighted_count = text_element.get('highlighted', 0)
        if highlighted_count > 0:
            segments.append(FormattedSegment(
                text=complete_text,
                start_position=0,
                end_position=len(complete_text),
                formatting_type=FormattingType.HIGHLIGHTED.value,
                formatting_details={'highlighted_count': highlighted_count},
                slide_number=slide_number,
                element_index=element_index
            ))
        
        return segments
    
    def _extract_strikethrough_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[FormattedSegment]:
        """Extract strikethrough text segments."""
        segments = []
        
        # Check if element has strikethrough formatting
        strikethrough_count = text_element.get('strikethrough', 0)
        if strikethrough_count > 0:
            segments.append(FormattedSegment(
                text=complete_text,
                start_position=0,
                end_position=len(complete_text),
                formatting_type=FormattingType.STRIKETHROUGH.value,
                formatting_details={'strikethrough_count': strikethrough_count},
                slide_number=slide_number,
                element_index=element_index
            ))
        
        return segments
    
    def _extract_hyperlink_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[HyperlinkSegment]:
        """Extract hyperlink text segments."""
        segments = []
        
        # Check if element has hyperlinks
        hyperlinks = text_element.get('hyperlinks', [])
        for hyperlink in hyperlinks:
            url = hyperlink.get('url', '')
            display_text = hyperlink.get('display_text', complete_text)
            
            # Determine link type
            link_type = "external"
            if url.startswith('mailto:'):
                link_type = "email"
            elif url.startswith('#'):
                link_type = "internal"
            
            # Calculate position of hyperlink text within complete text
            start_pos = complete_text.find(display_text)
            if start_pos == -1:
                start_pos = 0
            end_pos = start_pos + len(display_text)
            
            segments.append(HyperlinkSegment(
                text=display_text,
                start_position=start_pos,
                end_position=end_pos,
                formatting_type=FormattingType.HYPERLINKS.value,
                formatting_details={'url': url, 'link_type': link_type},
                slide_number=slide_number,
                element_index=element_index,
                url=url,
                display_text=display_text,
                link_type=link_type
            ))
        
        return segments
    
    def _extract_font_size_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[FormattedSegment]:
        """Extract text segments with font size information."""
        segments = []
        
        # Get font sizes from element
        font_sizes = text_element.get('font_sizes', [])
        if font_sizes:
            # For simplicity, use the first font size found
            font_size = font_sizes[0] if font_sizes else None
            if font_size:
                segments.append(FormattedSegment(
                    text=complete_text,
                    start_position=0,
                    end_position=len(complete_text),
                    formatting_type=FormattingType.FONT_SIZES.value,
                    formatting_details={'font_size': font_size, 'all_font_sizes': font_sizes},
                    slide_number=slide_number,
                    element_index=element_index
                ))
        
        return segments
    
    def _extract_font_color_segments(
        self, 
        text_element: Dict, 
        complete_text: str, 
        slide_number: int, 
        element_index: int
    ) -> List[FormattedSegment]:
        """Extract text segments with font color information."""
        segments = []
        
        # Get font colors from element
        font_colors = text_element.get('font_colors', [])
        if font_colors:
            # For simplicity, use the first font color found
            font_color = font_colors[0] if font_colors else None
            if font_color:
                segments.append(FormattedSegment(
                    text=complete_text,
                    start_position=0,
                    end_position=len(complete_text),
                    formatting_type=FormattingType.FONT_COLORS.value,
                    formatting_details={'font_color': font_color, 'all_font_colors': font_colors},
                    slide_number=slide_number,
                    element_index=element_index
                ))
        
        return segments
    
    def calculate_positions(
        self, 
        complete_text: str, 
        segments: List[str]
    ) -> List[Tuple[int, int]]:
        """Calculate start/end positions for formatted segments."""
        positions = []
        used_positions = set()  # Track used positions to handle overlapping
        
        for segment in segments:
            # Find all occurrences of this segment
            start_pos = 0
            found = False
            
            while True:
                pos = complete_text.find(segment, start_pos)
                if pos == -1:
                    break
                
                end_pos = pos + len(segment)
                
                # Check if this position range is already used
                position_range = set(range(pos, end_pos))
                if not position_range.intersection(used_positions):
                    positions.append((pos, end_pos))
                    used_positions.update(position_range)
                    found = True
                    break
                
                start_pos = pos + 1
            
            if not found:
                # If segment not found or all positions are used, add default position
                positions.append((0, len(segment)))
        
        return positions
    
    def handle_overlapping_formatting(
        self,
        text_element: Dict,
        complete_text: str,
        slide_number: int,
        element_index: int
    ) -> List[FormattedSegment]:
        """Handle overlapping formatting attributes on the same text."""
        segments = []
        
        # Check for multiple formatting types on the same element
        formatting_types = []
        
        if text_element.get('bolded', 0) > 0:
            formatting_types.append(FormattingType.BOLD.value)
        if text_element.get('italicized', 0) > 0:
            formatting_types.append(FormattingType.ITALIC.value)
        if text_element.get('underlined', 0) > 0:
            formatting_types.append(FormattingType.UNDERLINED.value)
        if text_element.get('highlighted', 0) > 0:
            formatting_types.append(FormattingType.HIGHLIGHTED.value)
        if text_element.get('strikethrough', 0) > 0:
            formatting_types.append(FormattingType.STRIKETHROUGH.value)
        
        # Create segments for each formatting type found
        for formatting_type in formatting_types:
            segment = FormattedSegment(
                text=complete_text,
                start_position=0,
                end_position=len(complete_text),
                formatting_type=formatting_type,
                formatting_details=self._get_formatting_details(text_element, formatting_type),
                slide_number=slide_number,
                element_index=element_index
            )
            segments.append(segment)
        
        return segments
    
    def _get_formatting_details(self, text_element: Dict, formatting_type: str) -> Dict[str, Any]:
        """Get formatting details for a specific formatting type."""
        details = {}
        
        if formatting_type == FormattingType.BOLD.value:
            details['bold_count'] = text_element.get('bolded', 0)
        elif formatting_type == FormattingType.ITALIC.value:
            details['italic_count'] = text_element.get('italicized', 0)
        elif formatting_type == FormattingType.UNDERLINED.value:
            details['underlined_count'] = text_element.get('underlined', 0)
        elif formatting_type == FormattingType.HIGHLIGHTED.value:
            details['highlighted_count'] = text_element.get('highlighted', 0)
        elif formatting_type == FormattingType.STRIKETHROUGH.value:
            details['strikethrough_count'] = text_element.get('strikethrough', 0)
        
        # Add common details
        details['font_sizes'] = text_element.get('font_sizes', [])
        details['font_colors'] = text_element.get('font_colors', [])
        
        return details
    
    def ensure_position_consistency(
        self,
        segments: List[FormattedSegment],
        encoding: str = 'utf-8'
    ) -> List[FormattedSegment]:
        """Ensure position consistency across different text encoding scenarios."""
        consistent_segments = []
        
        for segment in segments:
            # Ensure positions are valid
            text_length = len(segment.text)
            
            # Adjust positions if they're out of bounds
            start_pos = max(0, min(segment.start_position, text_length))
            end_pos = max(start_pos, min(segment.end_position, text_length))
            
            # Create new segment with consistent positions
            consistent_segment = FormattedSegment(
                text=segment.text,
                start_position=start_pos,
                end_position=end_pos,
                formatting_type=segment.formatting_type,
                formatting_details=segment.formatting_details,
                slide_number=segment.slide_number,
                element_index=segment.element_index
            )
            
            consistent_segments.append(consistent_segment)
        
        return consistent_segments
    
    def extract_hyperlink_details(
        self, 
        text_element: Dict
    ) -> List[HyperlinkSegment]:
        """Extract hyperlink text and URLs."""
        hyperlinks = []
        
        hyperlink_data = text_element.get('hyperlinks', [])
        for idx, hyperlink in enumerate(hyperlink_data):
            url = hyperlink.get('url', '')
            display_text = hyperlink.get('display_text', '')
            
            # Determine link type
            link_type = "external"
            if url.startswith('mailto:'):
                link_type = "email"
            elif url.startswith('#'):
                link_type = "internal"
            
            hyperlinks.append(HyperlinkSegment(
                text=display_text,
                start_position=0,  # Will be calculated later
                end_position=len(display_text),
                formatting_type=FormattingType.HYPERLINKS.value,
                formatting_details={'url': url, 'link_type': link_type},
                slide_number=0,  # Will be set by caller
                element_index=idx,
                url=url,
                display_text=display_text,
                link_type=link_type
            ))
        
        return hyperlinks