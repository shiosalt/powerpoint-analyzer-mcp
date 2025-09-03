"""
Analyze test_complex.pptx file to document its content and verify it meets test requirements.
"""

import sys
import json
from pathlib import Path

# Add the project root to the path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from powerpoint_mcp_server.core.content_extractor import ContentExtractor
from powerpoint_mcp_server.core.text_formatting_analyzer import TextFormattingAnalyzer
from powerpoint_mcp_server.core.enhanced_table_extractor import EnhancedTableExtractor
from powerpoint_mcp_server.utils.zip_extractor import ZipExtractor

def analyze_test_file(file_path: str):
    """Analyze the test file and document its content."""
    print(f"Analyzing test file: {file_path}")
    print("=" * 60)
    
    try:
        # Initialize extractors
        content_extractor = ContentExtractor()
        formatting_analyzer = TextFormattingAnalyzer(content_extractor)
        table_extractor = EnhancedTableExtractor(content_extractor)
        
        # Extract basic content
        with ZipExtractor(file_path) as extractor:
            # Get presentation metadata
            presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
            if presentation_xml:
                metadata = content_extractor.extract_presentation_metadata(presentation_xml)
                sections = content_extractor.extract_section_information(presentation_xml)
                
                print(f"Presentation Metadata:")
                print(f"  Title: {metadata.get('title', 'N/A')}")
                print(f"  Author: {metadata.get('author', 'N/A')}")
                print(f"  Slide Count: {metadata.get('slide_count', 0)}")
                print(f"  Sections: {len(sections)}")
                
                if sections:
                    print("  Section Details:")
                    for section in sections:
                        print(f"    - {section.get('name', 'Unnamed')}: slides {section.get('slide_range', [])}")
                print()
            
            # Get slide files
            slide_files_dict = extractor.get_slide_xml_files()
            slide_files = sorted(slide_files_dict.keys())
            
            print(f"Slide Analysis:")
            print(f"Total slides: {len(slide_files)}")
            print()
            
            # Analyze each slide
            for i, slide_file in enumerate(slide_files, 1):
                print(f"Slide {i}:")
                slide_xml = extractor.read_xml_content(slide_file)
                
                if slide_xml:
                    # Extract slide content
                    slide_info = content_extractor.extract_slide_content(slide_xml, i)
                    
                    print(f"  Title: {slide_info.title or 'No title'}")
                    print(f"  Subtitle: {slide_info.subtitle or 'No subtitle'}")
                    print(f"  Layout: {slide_info.layout_name}")
                    
                    # Count objects
                    root = content_extractor.xml_parser.parse_xml_string(slide_xml)
                    object_counts = content_extractor._count_slide_objects(root)
                    print(f"  Objects: {object_counts}")
                    
                    # Analyze formatting
                    from powerpoint_mcp_server.core.text_formatting_analyzer import FormattingFilter
                    formatting_filter = FormattingFilter()
                    formatting_elements = formatting_analyzer._extract_formatted_elements_from_slide(
                        slide_xml, i, formatting_filter
                    )
                    
                    if formatting_elements:
                        print(f"  Formatting found:")
                        for element in formatting_elements:
                            formatting = element.formatting
                            if formatting.get('has_formatting', False):
                                format_types = []
                                if formatting.get('bold_count', 0) > 0:
                                    format_types.append(f"bold({formatting['bold_count']})")
                                if formatting.get('italic_count', 0) > 0:
                                    format_types.append(f"italic({formatting['italic_count']})")
                                if formatting.get('underline_count', 0) > 0:
                                    format_types.append(f"underline({formatting['underline_count']})")
                                if formatting.get('highlight_count', 0) > 0:
                                    format_types.append(f"highlight({formatting['highlight_count']})")
                                if formatting.get('strikethrough_count', 0) > 0:
                                    format_types.append(f"strikethrough({formatting['strikethrough_count']})")
                                if formatting.get('font_colors'):
                                    format_types.append(f"colors({len(formatting['font_colors'])})")
                                if formatting.get('font_sizes'):
                                    format_types.append(f"sizes({len(formatting['font_sizes'])})")
                                if formatting.get('hyperlinks'):
                                    format_types.append(f"hyperlinks({len(formatting['hyperlinks'])})")
                                
                                if format_types:
                                    print(f"    - {element.content_type.value}: {', '.join(format_types)}")
                    
                    # Check for tables
                    from powerpoint_mcp_server.core.enhanced_table_extractor import TableCriteria, ColumnSelection, FormattingDetection
                    table_criteria = TableCriteria()
                    column_selection = ColumnSelection()
                    formatting_detection = FormattingDetection()
                    tables = table_extractor._extract_tables_from_slide(
                        slide_xml, i, table_criteria, column_selection, formatting_detection
                    )
                    
                    if tables:
                        print(f"  Tables found: {len(tables)}")
                        for j, table in enumerate(tables):
                            print(f"    - Table {j+1}: {table.rows}x{table.columns} ({', '.join(table.headers[:3])}{'...' if len(table.headers) > 3 else ''})")
                    
                    print()
        
        print("Analysis complete!")
        
        # Generate test requirements summary
        print("\nTest Requirements Summary:")
        print("=" * 40)
        
        # Check formatting requirements
        formatting_types_needed = [
            "bold", "italic", "underlined", "highlighted", "strikethrough", 
            "colored text", "hyperlinks", "font size variations"
        ]
        
        print("Required formatting types:")
        for fmt_type in formatting_types_needed:
            print(f"  - {fmt_type}: [Check manually in presentation]")
        
        print("\nRequired table content:")
        print("  - Simple table (3x3): [Check slide content]")
        print("  - Complex table (5x7 with formatting): [Check slide content]")
        print("  - Slides with no tables: [Check slide content]")
        
        print("\nRequired section structure:")
        print("  - Multiple sections with descriptive names: [Check presentation sections]")
        print("  - Slides distributed across sections: [Check section assignments]")
        
        print("\nRecommendations:")
        print("1. Ensure test_complex.pptx contains all required formatting types")
        print("2. Add tables to specific slides for table extraction testing")
        print("3. Create presentation sections if not present")
        print("4. Document expected results for each test scenario")
        
    except Exception as e:
        print(f"Error analyzing test file: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_file = "tests/test_files/test_complex.pptx"
    analyze_test_file(test_file)