"""
Simple table data extraction optimized for minimal context consumption.
"""

import logging
from typing import Dict, List, Any, Optional, Union
from html import escape

from .content_extractor import ContentExtractor
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)


class SimpleTableExtractor:
    """
    Simple table extractor optimized for minimal context consumption.
    Provides clean output formats without heavy formatting metadata.
    """

    def __init__(self, content_extractor: Optional[ContentExtractor] = None):
        """Initialize the simple table extractor."""
        self.content_extractor = content_extractor or ContentExtractor()

    def extract_tables_simple(
        self,
        file_path: str,
        slide_numbers: List[int],
        column_selection: Optional[Dict[str, Any]] = None,
        output_format: str = "row_col_value"
    ) -> Dict[str, Any]:
        """
        Extract table data in simplified format.

        Args:
            file_path: Path to the PowerPoint file
            slide_numbers: List of slide numbers to extract tables from
            column_selection: Optional column filtering configuration
            output_format: Output format ('row_col_value', 'row_col_formattedvalue', 'html', 'simple_html')

        Returns:
            Dictionary containing extracted table data in simplified format
        """
        logger.info(f"Extracting tables (simple) from slides {slide_numbers} in {file_path}")

        try:
            extracted_tables = []

            with ZipExtractor(file_path) as extractor:
                slide_files = extractor.get_slide_xml_files_sorted()
                total_slides = len(slide_files)

                # Validate slide numbers
                invalid_slides = [s for s in slide_numbers if s < 1 or s > total_slides]
                if invalid_slides:
                    raise ValueError(f"Invalid slide numbers: {invalid_slides}. Valid range: 1-{total_slides}")

                for slide_num in slide_numbers:
                    try:
                        slide_index = slide_num - 1
                        slide_file = slide_files[slide_index]
                        slide_xml = extractor.read_xml_content(slide_file)

                        if slide_xml:
                            tables = self._extract_tables_from_slide(
                                slide_xml, slide_num, column_selection, output_format
                            )
                            extracted_tables.extend(tables)

                    except Exception as slide_error:
                        logger.error(f"Error processing slide {slide_num}: {slide_error}")
                        continue

            # Format output based on requested format
            if output_format in ["html", "simple_html"]:
                return {"extracted_html_tables": extracted_tables}
            else:
                return {"extracted_tables": extracted_tables}

        except Exception as e:
            logger.error(f"Error extracting tables (simple): {e}")
            import traceback
            logger.error(f"Traceback: {traceback.format_exc()}")
            raise

    def _extract_tables_from_slide(
        self,
        slide_xml: str,
        slide_number: int,
        column_selection: Optional[Dict[str, Any]],
        output_format: str
    ) -> List[Dict[str, Any]]:
        """Extract tables from a single slide in simplified format."""
        try:
            root = self.content_extractor.xml_parser.parse_xml_string(slide_xml)
            if root is None:
                return []

            tables = []
            graphic_frames = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:graphicFrame'
            )

            for frame in graphic_frames:
                table = self._extract_simple_table_from_frame(
                    frame, slide_number, column_selection, output_format
                )
                if table:
                    tables.append(table)

            return tables

        except Exception as e:
            logger.warning(f"Failed to extract tables from slide {slide_number}: {e}")
            return []

    def _extract_simple_table_from_frame(
        self,
        graphic_frame,
        slide_number: int,
        column_selection: Optional[Dict[str, Any]],
        output_format: str
    ) -> Optional[Dict[str, Any]]:
        """Extract simple table data from a graphic frame."""
        try:
            table_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                graphic_frame, './/a:tbl'
            )

            if table_elem is None:
                return None

            # Parse table structure
            rows = self.content_extractor.xml_parser.find_elements_with_namespace(
                table_elem, './/a:tr'
            )

            if not rows:
                return None

            # Extract headers from first row
            first_row_cells = self.content_extractor.xml_parser.find_elements_with_namespace(
                rows[0], './/a:tc'
            )
            headers = []
            cells_to_skip = set()
            
            # First pass: identify cells to skip (those within merged cell spans)
            for col_index, cell_elem in enumerate(first_row_cells):
                col_span = int(cell_elem.get('gridSpan', '1'))
                if col_span > 1:
                    # Mark the next (col_span - 1) cells to skip
                    for skip_offset in range(1, col_span):
                        if col_index + skip_offset < len(first_row_cells):
                            cells_to_skip.add(col_index + skip_offset)
            
            # Second pass: build headers, skipping cells within merged spans
            for col_index, cell_elem in enumerate(first_row_cells):
                # Skip cells that are within a previous cell's span
                if col_index in cells_to_skip:
                    continue
                
                # Always use plain text for headers
                cell_text = self.content_extractor._extract_cell_text_content(cell_elem)
                headers.append(cell_text if cell_text.strip() else f"Column {len(headers) + 1}")

            # Apply column selection if specified
            if column_selection:
                headers = self._apply_column_filter(headers, column_selection)

            # Extract data based on output format
            if output_format in ["html", "simple_html"]:
                html_data = self._build_html_table(rows, headers, output_format)
                return {
                    "slide_number": slide_number,
                    "rows": len(rows),
                    "columns": len(headers),
                    "headers": headers,
                    "htmldata": html_data
                }
            else:
                # row_col_value or row_col_formattedvalue format
                data = self._build_row_col_data(rows, headers, output_format)
                return {
                    "slide_number": slide_number,
                    "rows": len(rows),
                    "columns": len(headers),
                    "headers": headers,
                    "data": data
                }

        except Exception as e:
            logger.warning(f"Failed to extract simple table from frame: {e}")
            return None

    def _extract_cell_text(self, cell_elem, output_format: str) -> str:
        """Extract text from a cell, optionally with formatting."""
        if output_format == "row_col_formattedvalue":
            # Return Markdown formatted text
            return self._extract_cell_text_with_markdown_formatting(cell_elem)
        else:
            # Plain text only (for row_col_value, html, simple_html)
            return self.content_extractor._extract_cell_text_content(cell_elem)

    def _extract_cell_text_with_markdown_formatting(self, cell_elem) -> str:
        """Extract cell text with inline Markdown formatting markers."""
        try:
            tx_body = self.content_extractor.xml_parser.find_element_with_namespace(
                cell_elem, './/a:txBody'
            )
            
            if tx_body is None:
                return ""
            
            # Get all paragraphs to preserve line breaks
            paragraphs = self.content_extractor.xml_parser.find_elements_with_namespace(
                tx_body, './/a:p'
            )
            
            if not paragraphs:
                return ""
            
            paragraph_texts = []
            
            for para in paragraphs:
                runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                    para, './/a:r'
                )
                
                para_text = ""
                
                for run in runs:
                    # Get text content
                    t_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        run, './/a:t'
                    )
                    if t_elem is None or t_elem.text is None:
                        continue
                    
                    # Normalize whitespace
                    text = ' '.join(t_elem.text.split())
                    if not text:
                        continue
                    
                    # Get run properties
                    r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                        run, './/a:rPr'
                    )
                    
                    if r_pr is not None:
                        # Check for bold - can be attribute or child element
                        is_bold = False
                        bold_attr = r_pr.get('b')
                        if bold_attr and bold_attr != '0':
                            is_bold = True
                        else:
                            bold_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:b'
                            )
                            if bold_elem is not None:
                                is_bold = bold_elem.get('val', '1') != '0'
                        
                        # Check for italic - can be attribute or child element
                        is_italic = False
                        italic_attr = r_pr.get('i')
                        if italic_attr and italic_attr != '0':
                            is_italic = True
                        else:
                            italic_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:i'
                            )
                            if italic_elem is not None:
                                is_italic = italic_elem.get('val', '1') != '0'
                        
                        # Check for underline - can be attribute or child element
                        is_underline = False
                        underline_attr = r_pr.get('u')
                        if underline_attr and underline_attr != 'none':
                            is_underline = True
                        else:
                            underline_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:u'
                            )
                            if underline_elem is not None:
                                is_underline = underline_elem.get('val', 'sng') != 'none'
                        
                        # Check for strikethrough - can be attribute or child element
                        is_strike = False
                        strike_attr = r_pr.get('strike')
                        if strike_attr and strike_attr != 'noStrike':
                            is_strike = True
                        else:
                            strike_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:strike'
                            )
                            if strike_elem is not None:
                                is_strike = strike_elem.get('val', 'sngStrike') != 'noStrike'
                        
                        # Apply Markdown formatting (order matters for proper nesting)
                        if is_bold:
                            text = f"**{text}**"
                        if is_italic:
                            text = f"*{text}*"
                        if is_underline:
                            text = f"_{text}_"
                        if is_strike:
                            text = f"~~{text}~~"
                    
                    para_text += text
                
                if para_text:
                    paragraph_texts.append(para_text)
            
            # Join paragraphs with newlines to preserve line breaks
            return "\n".join(paragraph_texts) if paragraph_texts else ""
            
        except Exception as e:
            logger.warning(f"Failed to extract cell text with Markdown formatting: {e}")
            import traceback
            logger.warning(f"Traceback: {traceback.format_exc()}")
            return self.content_extractor._extract_cell_text_content(cell_elem)

    def _extract_color_from_fill(self, fill_elem) -> Optional[str]:
        """Extract color value from a fill element."""
        try:
            # Look for RGB color
            srgb_clr = self.content_extractor.xml_parser.find_element_with_namespace(
                fill_elem, './/a:srgbClr'
            )
            if srgb_clr is not None:
                color_val = srgb_clr.get('val')
                if color_val:
                    return f"#{color_val}"
            
            # Look for scheme color
            scheme_clr = self.content_extractor.xml_parser.find_element_with_namespace(
                fill_elem, './/a:schemeClr'
            )
            if scheme_clr is not None:
                color_val = scheme_clr.get('val')
                if color_val:
                    return color_val
            
            return None
            
        except Exception as e:
            logger.warning(f"Failed to extract color from fill: {e}")
            return None

    def _build_row_col_data(
        self,
        rows,
        headers: List[str],
        output_format: str
    ) -> List[List[Union[int, str]]]:
        """Build row/col/value format data."""
        data = []

        for row_index, row_elem in enumerate(rows):
            cells = self.content_extractor.xml_parser.find_elements_with_namespace(
                row_elem, './/a:tc'
            )
            
            # Identify cells to skip in this row (those within merged cell spans)
            cells_to_skip = set()
            for col_index, cell_elem in enumerate(cells):
                col_span = int(cell_elem.get('gridSpan', '1'))
                if col_span > 1:
                    for skip_offset in range(1, col_span):
                        if col_index + skip_offset < len(cells):
                            cells_to_skip.add(col_index + skip_offset)
            
            # Build data, mapping physical cells to logical headers
            header_index = 0
            for col_index, cell_elem in enumerate(cells):
                # Skip cells within merged spans
                if col_index in cells_to_skip:
                    continue
                
                if header_index < len(headers):
                    cell_text = self._extract_cell_text(cell_elem, output_format)
                    # Format: [row, col, value]
                    data.append([row_index, header_index, cell_text])
                    header_index += 1

        return data

    def _build_html_table(
        self,
        rows,
        headers: List[str],
        output_format: str
    ) -> str:
        """Build HTML table with support for colspan/rowspan and formatting."""
        # Use different styling based on output format
        if output_format == "simple_html":
            html_parts = ['<table border="1" cellpadding="5" cellspacing="0" style="padding: 8px; white-space: pre-wrap;">']
        else:
            # Full HTML with formatting support
            html_parts = ['<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; padding: 8px; white-space: pre-wrap;">']

        # Add header row with thead styling
        html_parts.append('<thead style="background-color: #f0f0f0; font-weight: bold;"><tr>')
        for header in headers:
            html_parts.append(f'<th>{escape(header)}</th>')
        html_parts.append('</tr></thead>')

        # Add data rows
        html_parts.append('<tbody>')
        for row_elem in rows[1:]:  # Skip header row
            html_parts.append('<tr>')
            cells = self.content_extractor.xml_parser.find_elements_with_namespace(
                row_elem, './/a:tc'
            )
            
            # Identify cells to skip in this row (those within merged cell spans)
            cells_to_skip = set()
            for col_index, cell_elem in enumerate(cells):
                col_span = int(cell_elem.get('gridSpan', '1'))
                if col_span > 1:
                    for skip_offset in range(1, col_span):
                        if col_index + skip_offset < len(cells):
                            cells_to_skip.add(col_index + skip_offset)
            
            # Build HTML cells, mapping physical cells to logical headers
            header_index = 0
            for col_index, cell_elem in enumerate(cells):
                # Skip cells within merged spans
                if col_index in cells_to_skip:
                    continue
                
                if header_index < len(headers):
                    # Get rowspan and colspan
                    row_span = int(cell_elem.get('rowSpan', '1'))
                    col_span = int(cell_elem.get('gridSpan', '1'))

                    # Get cell background color (only add style if there's a custom background)
                    cell_styles = []
                    tc_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                        cell_elem, './/a:tcPr'
                    )
                    if tc_pr is not None:
                        solid_fill = self.content_extractor.xml_parser.find_element_with_namespace(
                            tc_pr, './/a:solidFill'
                        )
                        if solid_fill is not None:
                            bg_color = self._extract_color_from_fill(solid_fill)
                            if bg_color:
                                cell_styles.append(f'background-color: {bg_color}')

                    # Extract cell content with formatting
                    if output_format == "html":
                        cell_html = self._extract_cell_html_with_formatting(cell_elem)
                    else:
                        cell_text = self._extract_cell_text(cell_elem, "row_col_value")
                        cell_html = escape(cell_text).replace('\n', '<br>')

                    # Build cell tag with span attributes
                    cell_attrs = []
                    if cell_styles:
                        cell_attrs.append(f'style="{"; ".join(cell_styles)}"')
                    if row_span > 1:
                        cell_attrs.append(f'rowspan="{row_span}"')
                    if col_span > 1:
                        cell_attrs.append(f'colspan="{col_span}"')

                    attrs_str = ' ' + ' '.join(cell_attrs) if cell_attrs else ''
                    html_parts.append(f'<td{attrs_str}>{cell_html}</td>')
                    header_index += 1

            html_parts.append('</tr>')
        html_parts.append('</tbody>')

        html_parts.append('</table>')
        return ''.join(html_parts)
    
    def _extract_cell_html_with_formatting(self, cell_elem) -> str:
        """Extract cell content as HTML with inline formatting."""
        try:
            # Find text body
            tx_body = self.content_extractor.xml_parser.find_element_with_namespace(
                cell_elem, './/a:txBody'
            )
            
            if tx_body is None:
                return ""
            
            # Get all paragraphs to preserve line breaks
            paragraphs = self.content_extractor.xml_parser.find_elements_with_namespace(
                tx_body, './/a:p'
            )
            
            if not paragraphs:
                return ""
            
            html_parts = []
            
            for para_idx, para in enumerate(paragraphs):
                # Find all runs (text segments with formatting)
                runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                    para, './/a:r'
                )
                
                for run in runs:
                    # Get text content
                    t_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                        run, './/a:t'
                    )
                    if t_elem is None or t_elem.text is None:
                        continue
                    
                    text = escape(t_elem.text)
                    
                    # Get run properties
                    r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                        run, './/a:rPr'
                    )
                    
                    if r_pr is not None:
                        styles = []
                        tags_open = []
                        tags_close = []
                        
                        # Check for bold - can be attribute or child element
                        is_bold = False
                        bold_attr = r_pr.get('b')
                        if bold_attr and bold_attr != '0':
                            is_bold = True
                        else:
                            bold_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:b'
                            )
                            if bold_elem is not None:
                                is_bold = bold_elem.get('val', '1') != '0'
                        
                        if is_bold:
                            tags_open.append('<strong>')
                            tags_close.insert(0, '</strong>')
                        
                        # Check for italic - can be attribute or child element
                        is_italic = False
                        italic_attr = r_pr.get('i')
                        if italic_attr and italic_attr != '0':
                            is_italic = True
                        else:
                            italic_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:i'
                            )
                            if italic_elem is not None:
                                is_italic = italic_elem.get('val', '1') != '0'
                        
                        if is_italic:
                            tags_open.append('<em>')
                            tags_close.insert(0, '</em>')
                        
                        # Check for underline - can be attribute or child element
                        is_underline = False
                        underline_attr = r_pr.get('u')
                        if underline_attr and underline_attr != 'none':
                            is_underline = True
                        else:
                            underline_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:u'
                            )
                            if underline_elem is not None:
                                is_underline = underline_elem.get('val', 'sng') != 'none'
                        
                        if is_underline:
                            tags_open.append('<u>')
                            tags_close.insert(0, '</u>')
                        
                        # Check for strikethrough - can be attribute or child element
                        is_strike = False
                        strike_attr = r_pr.get('strike')
                        if strike_attr and strike_attr != 'noStrike':
                            is_strike = True
                        else:
                            strike_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:strike'
                            )
                            if strike_elem is not None:
                                is_strike = strike_elem.get('val', 'sngStrike') != 'noStrike'
                        
                        if is_strike:
                            tags_open.append('<s>')
                            tags_close.insert(0, '</s>')
                        
                        # Get font color
                        solid_fill = self.content_extractor.xml_parser.find_element_with_namespace(
                            r_pr, './/a:solidFill'
                        )
                        if solid_fill is not None:
                            font_color = self._extract_color_from_fill(solid_fill)
                            if font_color:
                                styles.append(f'color: {font_color}')
                        
                        # Get highlight color
                        highlight_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                            r_pr, './/a:highlight'
                        )
                        if highlight_elem is not None:
                            highlight_color = self._extract_color_from_fill(highlight_elem)
                            if highlight_color:
                                styles.append(f'background-color: {highlight_color}')
                        
                        # Check for hyperlinks
                        hlinkClick = self.content_extractor.xml_parser.find_element_with_namespace(
                            r_pr, './/a:hlinkClick'
                        )
                        if hlinkClick is not None:
                            r_id = hlinkClick.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                            if r_id:
                                tags_open.insert(0, f'<a href="#{r_id}">')
                                tags_close.append('</a>')
                        
                        # Build HTML with formatting
                        if styles:
                            html_parts.append(f'<span style="{"; ".join(styles)}">')
                        
                        html_parts.extend(tags_open)
                        html_parts.append(text)
                        html_parts.extend(tags_close)
                        
                        if styles:
                            html_parts.append('</span>')
                    else:
                        html_parts.append(text)
                
                # Add line break after each paragraph except the last
                if para_idx < len(paragraphs) - 1:
                    html_parts.append('<br>')
            
            return ''.join(html_parts)
            
        except Exception as e:
            logger.warning(f"Failed to extract HTML formatted cell: {e}")
            import traceback
            logger.warning(f"Traceback: {traceback.format_exc()}")
            return escape(self.content_extractor._extract_cell_text_content(cell_elem))

    def _apply_column_filter(
        self,
        headers: List[str],
        column_selection: Dict[str, Any]
    ) -> List[str]:
        """Apply column selection filter to headers."""
        if not column_selection:
            return headers

        specific_columns = column_selection.get('specific_columns')
        exclude_columns = column_selection.get('exclude_columns')

        if specific_columns:
            # Include only specific columns
            filtered = []
            for col_name in specific_columns:
                for header in headers:
                    if col_name.lower() == header.lower():
                        filtered.append(header)
                        break
            return filtered

        if exclude_columns:
            # Exclude specific columns
            return [h for h in headers if h.lower() not in [e.lower() for e in exclude_columns]]

        return headers
