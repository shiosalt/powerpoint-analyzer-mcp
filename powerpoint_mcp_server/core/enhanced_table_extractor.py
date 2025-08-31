"""
Enhanced table data extraction with flexible selection and formatting detection.
"""

import re
import logging
from typing import Dict, List, Any, Optional, Union, Tuple
from dataclasses import dataclass, field
from enum import Enum

from .content_extractor import ContentExtractor
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)


class OutputFormat(Enum):
    """Enumeration of available output formats."""
    STRUCTURED = "structured"
    FLAT = "flat"
    GROUPED_BY_SLIDE = "grouped_by_slide"


@dataclass
class TableCriteria:
    """Criteria for selecting tables."""
    min_rows: Optional[int] = None
    min_columns: Optional[int] = None
    max_rows: Optional[int] = None
    max_columns: Optional[int] = None
    header_contains: Optional[List[str]] = None
    header_patterns: Optional[List[str]] = None


@dataclass
class ColumnSelection:
    """Configuration for column selection."""
    specific_columns: Optional[List[str]] = None
    column_patterns: Optional[List[str]] = None
    exclude_columns: Optional[List[str]] = None
    all_columns: bool = True


@dataclass
class FormattingDetection:
    """Configuration for formatting detection."""
    detect_bold: bool = True
    detect_italic: bool = True
    detect_underline: bool = True
    detect_highlight: bool = True
    detect_colors: bool = True
    detect_hyperlinks: bool = True
    preserve_formatting: bool = False


@dataclass
class CellFormatting:
    """Formatting information for a table cell."""
    bold: bool = False
    italic: bool = False
    underline: bool = False
    highlight: bool = False
    strikethrough: bool = False
    font_color: Optional[str] = None
    background_color: Optional[str] = None
    font_size: Optional[int] = None
    hyperlink: Optional[str] = None


@dataclass
class EnhancedTableCell:
    """Enhanced table cell with formatting information."""
    value: str
    formatting: CellFormatting = field(default_factory=CellFormatting)
    row_span: int = 1
    col_span: int = 1
    position: Tuple[int, int] = (0, 0)


@dataclass
class EnhancedTable:
    """Enhanced table with metadata and formatting."""
    slide_number: int
    table_index: int
    rows: int
    columns: int
    headers: List[str]
    data: List[Dict[str, EnhancedTableCell]]
    metadata: Dict[str, Any] = field(default_factory=dict)
    position: Tuple[int, int] = (0, 0)
    size: Tuple[int, int] = (0, 0)


class EnhancedTableExtractor:
    """
    Enhanced table extractor with flexible selection and formatting detection.
    """
    
    def __init__(self, content_extractor: Optional[ContentExtractor] = None):
        """Initialize the enhanced table extractor."""
        self.content_extractor = content_extractor or ContentExtractor()
        self._table_cache = {}
    
    def extract_tables(
        self,
        file_path: str,
        slide_numbers: List[int],
        table_criteria: Optional[TableCriteria] = None,
        column_selection: Optional[ColumnSelection] = None,
        formatting_detection: Optional[FormattingDetection] = None,
        output_format: OutputFormat = OutputFormat.STRUCTURED,
        include_metadata: bool = True
    ) -> Dict[str, Any]:
        """
        Extract table data with flexible selection and formatting.
        
        Args:
            file_path: Path to the PowerPoint file
            slide_numbers: List of slide numbers to extract tables from
            table_criteria: Criteria for selecting tables
            column_selection: Configuration for column selection
            formatting_detection: Configuration for formatting detection
            output_format: Output format for extracted data
            include_metadata: Whether to include table metadata
            
        Returns:
            Dictionary containing extracted table data
        """
        if table_criteria is None:
            table_criteria = TableCriteria()
        if column_selection is None:
            column_selection = ColumnSelection()
        if formatting_detection is None:
            formatting_detection = FormattingDetection()
        
        logger.info(f"Extracting tables from slides {slide_numbers} in {file_path}")
        
        try:
            extracted_tables = []
            
            with ZipExtractor(file_path) as extractor:
                # Get slide XML files
                slide_files = extractor.get_slide_xml_files()
                
                for slide_num in slide_numbers:
                    if slide_num <= len(slide_files):
                        slide_file = slide_files[slide_num - 1]
                        slide_xml = extractor.read_xml_content(slide_file)
                        
                        if slide_xml:
                            tables = self._extract_tables_from_slide(
                                slide_xml, slide_num, table_criteria,
                                column_selection, formatting_detection
                            )
                            extracted_tables.extend(tables)
            
            # Format output based on requested format
            result = self._format_output(
                extracted_tables, output_format, include_metadata
            )
            
            logger.info(f"Extracted {len(extracted_tables)} tables")
            return result
            
        except Exception as e:
            logger.error(f"Error extracting tables: {e}")
            raise
    
    def _extract_tables_from_slide(
        self,
        slide_xml: str,
        slide_number: int,
        table_criteria: TableCriteria,
        column_selection: ColumnSelection,
        formatting_detection: FormattingDetection
    ) -> List[EnhancedTable]:
        """Extract tables from a single slide with enhanced processing."""
        try:
            root = self.content_extractor.xml_parser.parse_xml_string(slide_xml)
            if root is None:
                return []
            
            tables = []
            
            # Find all graphic frames that might contain tables
            graphic_frames = self.content_extractor.xml_parser.find_elements_with_namespace(
                root, './/p:graphicFrame'
            )
            
            for table_index, frame in enumerate(graphic_frames):
                table = self._extract_enhanced_table_from_frame(
                    frame, slide_number, table_index, table_criteria,
                    column_selection, formatting_detection
                )
                
                if table and self._meets_table_criteria(table, table_criteria):
                    tables.append(table)
            
            return tables
            
        except Exception as e:
            logger.warning(f"Failed to extract tables from slide {slide_number}: {e}")
            return []
    
    def _extract_enhanced_table_from_frame(
        self,
        graphic_frame,
        slide_number: int,
        table_index: int,
        table_criteria: TableCriteria,
        column_selection: ColumnSelection,
        formatting_detection: FormattingDetection
    ) -> Optional[EnhancedTable]:
        """Extract enhanced table data from a graphic frame."""
        try:
            # Check if this graphic frame contains a table
            table_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                graphic_frame, './/a:tbl'
            )
            
            if table_elem is None:
                return None
            
            # Extract position and size
            position, size = self.content_extractor._extract_graphic_frame_transform(graphic_frame)
            
            # Parse table structure with enhanced formatting
            enhanced_table = self._parse_enhanced_table_structure(
                table_elem, slide_number, table_index, formatting_detection
            )
            
            if enhanced_table is None:
                return None
            
            # Set position and size
            enhanced_table.position = position
            enhanced_table.size = size
            
            # Apply column selection
            if not column_selection.all_columns:
                enhanced_table = self._apply_column_selection(enhanced_table, column_selection)
            
            return enhanced_table
            
        except Exception as e:
            logger.warning(f"Failed to extract enhanced table from frame: {e}")
            return None
    
    def _parse_enhanced_table_structure(
        self,
        table_elem,
        slide_number: int,
        table_index: int,
        formatting_detection: FormattingDetection
    ) -> Optional[EnhancedTable]:
        """Parse table structure with enhanced formatting detection."""
        try:
            # Find all table rows
            rows = self.content_extractor.xml_parser.find_elements_with_namespace(
                table_elem, './/a:tr'
            )
            
            if not rows:
                return None
            
            table_data = []
            headers = []
            max_columns = 0
            
            for row_index, row_elem in enumerate(rows):
                # Find all cells in this row
                cells = self.content_extractor.xml_parser.find_elements_with_namespace(
                    row_elem, './/a:tc'
                )
                
                row_data = {}
                
                for col_index, cell_elem in enumerate(cells):
                    cell = self._parse_enhanced_table_cell(
                        cell_elem, formatting_detection, row_index, col_index
                    )
                    
                    # Use column index as key for now, will be replaced with headers
                    column_key = f"col_{col_index}"
                    row_data[column_key] = cell
                    
                    # Extract headers from first row
                    if row_index == 0:
                        headers.append(cell.value if cell.value.strip() else f"Column {col_index + 1}")
                
                table_data.append(row_data)
                max_columns = max(max_columns, len(cells))
            
            # Replace column keys with actual headers
            formatted_data = []
            for row_data in table_data:
                formatted_row = {}
                for col_index in range(max_columns):
                    old_key = f"col_{col_index}"
                    if old_key in row_data:
                        header = headers[col_index] if col_index < len(headers) else f"Column {col_index + 1}"
                        formatted_row[header] = row_data[old_key]
                formatted_data.append(formatted_row)
            
            # Create enhanced table
            enhanced_table = EnhancedTable(
                slide_number=slide_number,
                table_index=table_index,
                rows=len(formatted_data),
                columns=max_columns,
                headers=headers,
                data=formatted_data,
                metadata={
                    'has_formatting': self._has_formatting(formatted_data),
                    'cell_count': len(formatted_data) * max_columns,
                    'non_empty_cells': self._count_non_empty_cells(formatted_data)
                }
            )
            
            return enhanced_table
            
        except Exception as e:
            logger.warning(f"Failed to parse enhanced table structure: {e}")
            return None
    
    def _parse_enhanced_table_cell(
        self,
        cell_elem,
        formatting_detection: FormattingDetection,
        row_index: int,
        col_index: int
    ) -> EnhancedTableCell:
        """Parse a single table cell with enhanced formatting detection."""
        try:
            # Extract basic cell content
            content = self.content_extractor._extract_cell_text_content(cell_elem)
            
            # Extract row span and column span
            row_span = int(cell_elem.get('rowSpan', '1'))
            col_span = int(cell_elem.get('gridSpan', '1'))
            
            # Initialize formatting
            formatting = CellFormatting()
            
            if formatting_detection.detect_bold or formatting_detection.detect_italic or \
               formatting_detection.detect_underline or formatting_detection.detect_colors:
                formatting = self._extract_enhanced_cell_formatting(
                    cell_elem, formatting_detection
                )
            
            return EnhancedTableCell(
                value=content,
                formatting=formatting,
                row_span=row_span,
                col_span=col_span,
                position=(row_index, col_index)
            )
            
        except Exception as e:
            logger.warning(f"Failed to parse enhanced table cell: {e}")
            return EnhancedTableCell(value="", position=(row_index, col_index))
    
    def _extract_enhanced_cell_formatting(
        self,
        cell_elem,
        formatting_detection: FormattingDetection
    ) -> CellFormatting:
        """Extract enhanced formatting information from a table cell."""
        try:
            formatting = CellFormatting()
            
            # Extract cell background color
            if formatting_detection.detect_colors:
                tc_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                    cell_elem, './/a:tcPr'
                )
                if tc_pr is not None:
                    solid_fill = self.content_extractor.xml_parser.find_element_with_namespace(
                        tc_pr, './/a:solidFill'
                    )
                    if solid_fill is not None:
                        color = self._extract_color_from_fill(solid_fill)
                        if color:
                            formatting.background_color = color
            
            # Extract text formatting from runs
            tx_body = self.content_extractor.xml_parser.find_element_with_namespace(
                cell_elem, './/a:txBody'
            )
            
            if tx_body is not None:
                runs = self.content_extractor.xml_parser.find_elements_with_namespace(
                    tx_body, './/a:r'
                )
                
                for run in runs:
                    r_pr = self.content_extractor.xml_parser.find_element_with_namespace(
                        run, './/a:rPr'
                    )
                    
                    if r_pr is not None:
                        # Check for bold
                        if formatting_detection.detect_bold:
                            bold_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:b'
                            )
                            if bold_elem is not None and bold_elem.get('val', '1') != '0':
                                formatting.bold = True
                        
                        # Check for italic
                        if formatting_detection.detect_italic:
                            italic_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:i'
                            )
                            if italic_elem is not None and italic_elem.get('val', '1') != '0':
                                formatting.italic = True
                        
                        # Check for underline
                        if formatting_detection.detect_underline:
                            underline_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:u'
                            )
                            if underline_elem is not None and underline_elem.get('val', 'sng') != 'none':
                                formatting.underline = True
                        
                        # Check for strikethrough
                        strike_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                            r_pr, './/a:strike'
                        )
                        if strike_elem is not None and strike_elem.get('val', 'sngStrike') != 'noStrike':
                            formatting.strikethrough = True
                        
                        # Extract font color
                        if formatting_detection.detect_colors:
                            solid_fill = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:solidFill'
                            )
                            if solid_fill is not None:
                                color = self._extract_color_from_fill(solid_fill)
                                if color:
                                    formatting.font_color = color
                        
                        # Extract font size
                        font_size_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                            r_pr, './/a:sz'
                        )
                        if font_size_elem is not None:
                            sz = font_size_elem.get('val')
                            if sz:
                                # Font size in PowerPoint is in hundredths of a point
                                formatting.font_size = int(sz) // 100
                        
                        # Check for highlight
                        if formatting_detection.detect_highlight:
                            highlight_elem = self.content_extractor.xml_parser.find_element_with_namespace(
                                r_pr, './/a:highlight'
                            )
                            if highlight_elem is not None:
                                formatting.highlight = True
                
                # Check for hyperlinks
                if formatting_detection.detect_hyperlinks:
                    hyperlinks = self.content_extractor.xml_parser.find_elements_with_namespace(
                        tx_body, './/a:hlinkClick'
                    )
                    if hyperlinks:
                        # For now, just mark that there's a hyperlink
                        # In a full implementation, we'd resolve the relationship ID
                        formatting.hyperlink = "present"
            
            return formatting
            
        except Exception as e:
            logger.warning(f"Failed to extract enhanced cell formatting: {e}")
            return CellFormatting()
    
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
    
    def _meets_table_criteria(self, table: EnhancedTable, criteria: TableCriteria) -> bool:
        """Check if a table meets the specified criteria."""
        try:
            # Check row count
            if criteria.min_rows is not None and table.rows < criteria.min_rows:
                return False
            if criteria.max_rows is not None and table.rows > criteria.max_rows:
                return False
            
            # Check column count
            if criteria.min_columns is not None and table.columns < criteria.min_columns:
                return False
            if criteria.max_columns is not None and table.columns > criteria.max_columns:
                return False
            
            # Check header contains
            if criteria.header_contains:
                for required_header in criteria.header_contains:
                    if not any(required_header.lower() in header.lower() for header in table.headers):
                        return False
            
            # Check header patterns
            if criteria.header_patterns:
                for pattern in criteria.header_patterns:
                    try:
                        if not any(re.search(pattern, header, re.IGNORECASE) for header in table.headers):
                            return False
                    except re.error:
                        # Fallback to simple string matching
                        if not any(pattern.lower() in header.lower() for header in table.headers):
                            return False
            
            return True
            
        except Exception as e:
            logger.warning(f"Failed to check table criteria: {e}")
            return True  # Default to including the table
    
    def _apply_column_selection(
        self, 
        table: EnhancedTable, 
        column_selection: ColumnSelection
    ) -> EnhancedTable:
        """Apply column selection to filter table columns."""
        try:
            selected_headers = []
            
            # Determine which columns to include
            if column_selection.specific_columns:
                # Include specific columns by name
                for col_name in column_selection.specific_columns:
                    # Find matching headers (case-insensitive)
                    for header in table.headers:
                        if col_name.lower() == header.lower():
                            selected_headers.append(header)
                            break
            
            elif column_selection.column_patterns:
                # Include columns matching patterns
                for pattern in column_selection.column_patterns:
                    try:
                        for header in table.headers:
                            if re.search(pattern, header, re.IGNORECASE):
                                if header not in selected_headers:
                                    selected_headers.append(header)
                    except re.error:
                        # Fallback to simple string matching
                        for header in table.headers:
                            if pattern.lower() in header.lower():
                                if header not in selected_headers:
                                    selected_headers.append(header)
            
            else:
                # Start with all headers
                selected_headers = table.headers.copy()
            
            # Remove excluded columns
            if column_selection.exclude_columns:
                for exclude_col in column_selection.exclude_columns:
                    selected_headers = [h for h in selected_headers if h.lower() != exclude_col.lower()]
            
            # Filter table data
            filtered_data = []
            for row in table.data:
                filtered_row = {}
                for header in selected_headers:
                    if header in row:
                        filtered_row[header] = row[header]
                filtered_data.append(filtered_row)
            
            # Update table
            table.headers = selected_headers
            table.columns = len(selected_headers)
            table.data = filtered_data
            
            return table
            
        except Exception as e:
            logger.warning(f"Failed to apply column selection: {e}")
            return table
    
    def _has_formatting(self, table_data: List[Dict[str, EnhancedTableCell]]) -> bool:
        """Check if any cells in the table have formatting."""
        try:
            for row in table_data:
                for cell in row.values():
                    formatting = cell.formatting
                    if (formatting.bold or formatting.italic or formatting.underline or
                        formatting.highlight or formatting.strikethrough or
                        formatting.font_color or formatting.background_color or
                        formatting.hyperlink):
                        return True
            return False
        except Exception:
            return False
    
    def _count_non_empty_cells(self, table_data: List[Dict[str, EnhancedTableCell]]) -> int:
        """Count non-empty cells in the table."""
        try:
            count = 0
            for row in table_data:
                for cell in row.values():
                    if cell.value.strip():
                        count += 1
            return count
        except Exception:
            return 0
    
    def _format_output(
        self,
        tables: List[EnhancedTable],
        output_format: OutputFormat,
        include_metadata: bool
    ) -> Dict[str, Any]:
        """Format the output based on the requested format."""
        try:
            if output_format == OutputFormat.STRUCTURED:
                return self._format_structured_output(tables, include_metadata)
            elif output_format == OutputFormat.FLAT:
                return self._format_flat_output(tables, include_metadata)
            elif output_format == OutputFormat.GROUPED_BY_SLIDE:
                return self._format_grouped_output(tables, include_metadata)
            else:
                return self._format_structured_output(tables, include_metadata)
                
        except Exception as e:
            logger.warning(f"Failed to format output: {e}")
            return {"extracted_tables": [], "summary": {"total_tables": 0}}
    
    def _format_structured_output(
        self, 
        tables: List[EnhancedTable], 
        include_metadata: bool
    ) -> Dict[str, Any]:
        """Format output in structured format."""
        extracted_tables = []
        
        for table in tables:
            table_dict = {
                "slide_number": table.slide_number,
                "table_index": table.table_index,
                "rows": table.rows,
                "columns": table.columns,
                "headers": table.headers,
                "data": []
            }
            
            if include_metadata:
                table_dict["metadata"] = table.metadata
                table_dict["position"] = table.position
                table_dict["size"] = table.size
            
            # Convert data to serializable format
            for row in table.data:
                row_dict = {}
                for header, cell in row.items():
                    cell_dict = {
                        "value": cell.value,
                        "formatting": {
                            "bold": cell.formatting.bold,
                            "italic": cell.formatting.italic,
                            "underline": cell.formatting.underline,
                            "highlight": cell.formatting.highlight,
                            "strikethrough": cell.formatting.strikethrough,
                            "font_color": cell.formatting.font_color,
                            "background_color": cell.formatting.background_color,
                            "font_size": cell.formatting.font_size,
                            "hyperlink": cell.formatting.hyperlink
                        }
                    }
                    
                    if include_metadata:
                        cell_dict["row_span"] = cell.row_span
                        cell_dict["col_span"] = cell.col_span
                        cell_dict["position"] = cell.position
                    
                    row_dict[header] = cell_dict
                
                table_dict["data"].append(row_dict)
            
            extracted_tables.append(table_dict)
        
        # Create summary
        summary = {
            "total_tables": len(tables),
            "total_rows": sum(table.rows for table in tables),
            "slides_with_tables": len(set(table.slide_number for table in tables)),
            "formatting_found": {
                "bold_cells": self._count_formatted_cells(tables, "bold"),
                "italic_cells": self._count_formatted_cells(tables, "italic"),
                "highlighted_cells": self._count_formatted_cells(tables, "highlight"),
                "colored_cells": self._count_formatted_cells(tables, "color")
            }
        }
        
        return {
            "extracted_tables": extracted_tables,
            "summary": summary
        }
    
    def _format_flat_output(
        self, 
        tables: List[EnhancedTable], 
        include_metadata: bool
    ) -> Dict[str, Any]:
        """Format output in flat format (all rows from all tables)."""
        all_rows = []
        
        for table in tables:
            for row_index, row in enumerate(table.data):
                flat_row = {
                    "slide_number": table.slide_number,
                    "table_index": table.table_index,
                    "row_index": row_index
                }
                
                for header, cell in row.items():
                    flat_row[header] = cell.value
                    
                    # Add formatting info if requested
                    if include_metadata:
                        flat_row[f"{header}_bold"] = cell.formatting.bold
                        flat_row[f"{header}_italic"] = cell.formatting.italic
                        flat_row[f"{header}_highlight"] = cell.formatting.highlight
                        if cell.formatting.font_color:
                            flat_row[f"{header}_color"] = cell.formatting.font_color
                
                all_rows.append(flat_row)
        
        return {
            "data": all_rows,
            "summary": {
                "total_rows": len(all_rows),
                "total_tables": len(tables)
            }
        }
    
    def _format_grouped_output(
        self, 
        tables: List[EnhancedTable], 
        include_metadata: bool
    ) -> Dict[str, Any]:
        """Format output grouped by slide."""
        slides = {}
        
        for table in tables:
            slide_num = table.slide_number
            if slide_num not in slides:
                slides[slide_num] = {
                    "slide_number": slide_num,
                    "tables": []
                }
            
            table_dict = self._format_structured_output([table], include_metadata)
            slides[slide_num]["tables"].extend(table_dict["extracted_tables"])
        
        return {
            "slides": list(slides.values()),
            "summary": {
                "total_slides": len(slides),
                "total_tables": len(tables)
            }
        }
    
    def _count_formatted_cells(self, tables: List[EnhancedTable], format_type: str) -> int:
        """Count cells with specific formatting across all tables."""
        count = 0
        for table in tables:
            for row in table.data:
                for cell in row.values():
                    if format_type == "bold" and cell.formatting.bold:
                        count += 1
                    elif format_type == "italic" and cell.formatting.italic:
                        count += 1
                    elif format_type == "highlight" and cell.formatting.highlight:
                        count += 1
                    elif format_type == "color" and (cell.formatting.font_color or cell.formatting.background_color):
                        count += 1
        return count
    
    def clear_cache(self):
        """Clear the table extraction cache."""
        self._table_cache.clear()
        logger.debug("Table extraction cache cleared")


def create_table_criteria_from_dict(criteria_dict: Dict[str, Any]) -> TableCriteria:
    """Create TableCriteria from a dictionary representation."""
    return TableCriteria(
        min_rows=criteria_dict.get('min_rows'),
        min_columns=criteria_dict.get('min_columns'),
        max_rows=criteria_dict.get('max_rows'),
        max_columns=criteria_dict.get('max_columns'),
        header_contains=criteria_dict.get('header_contains'),
        header_patterns=criteria_dict.get('header_patterns')
    )


def create_column_selection_from_dict(selection_dict: Dict[str, Any]) -> ColumnSelection:
    """Create ColumnSelection from a dictionary representation."""
    return ColumnSelection(
        specific_columns=selection_dict.get('specific_columns'),
        column_patterns=selection_dict.get('column_patterns'),
        exclude_columns=selection_dict.get('exclude_columns'),
        all_columns=selection_dict.get('all_columns', True)
    )


def create_formatting_detection_from_dict(detection_dict: Dict[str, Any]) -> FormattingDetection:
    """Create FormattingDetection from a dictionary representation."""
    return FormattingDetection(
        detect_bold=detection_dict.get('detect_bold', True),
        detect_italic=detection_dict.get('detect_italic', True),
        detect_underline=detection_dict.get('detect_underline', True),
        detect_highlight=detection_dict.get('detect_highlight', True),
        detect_colors=detection_dict.get('detect_colors', True),
        detect_hyperlinks=detection_dict.get('detect_hyperlinks', True),
        preserve_formatting=detection_dict.get('preserve_formatting', False)
    )