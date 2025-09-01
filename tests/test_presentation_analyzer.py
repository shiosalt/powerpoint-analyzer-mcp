"""
Unit tests for PresentationAnalyzer.
"""

import pytest
from unittest.mock import Mock, patch, MagicMock

from powerpoint_mcp_server.core.presentation_analyzer import (
    PresentationAnalyzer,
    AnalysisDepth,
    SlideType,
    SlideClassification,
    ContentPattern,
    PresentationStructure,
    PresentationInsights,
    PresentationOverview
)


class TestPresentationAnalyzer:
    """Test cases for PresentationAnalyzer."""
    
    @pytest.fixture
    def mock_content_extractor(self):
        """Create a mock content extractor."""
        extractor = Mock()
        
        # Mock XML parser
        xml_parser = Mock()
        xml_parser.parse_xml_string.return_value = Mock()
        extractor.xml_parser = xml_parser
        
        # Mock methods
        extractor.extract_presentation_metadata.return_value = {
            'slide_count': 5,
            'slide_size': {'width': 10, 'height': 7.5}
        }
        extractor.extract_section_information.return_value = [
            {'name': 'Introduction', 'id': '1'},
            {'name': 'Main Content', 'id': '2'}
        ]
        extractor._count_slide_objects.return_value = {
            'shapes': 2, 'text_boxes': 1, 'images': 0, 'tables': 0, 'charts': 0
        }
        extractor._extract_notes_content.return_value = "Sample notes"
        
        return extractor
    
    @pytest.fixture
    def presentation_analyzer(self, mock_content_extractor):
        """Create a PresentationAnalyzer with mocked dependencies."""
        return PresentationAnalyzer(mock_content_extractor)
    
    @pytest.fixture
    def sample_slide_info(self):
        """Create sample slide info."""
        slide_info = Mock()
        slide_info.title = "Sample Title"
        slide_info.subtitle = "Sample Subtitle"
        slide_info.text_elements = [
            {'content_plain': 'First text element'},
            {'content_plain': 'Second text element'}
        ]
        slide_info.tables = []
        slide_info.placeholders = []
        return slide_info
    
    @pytest.fixture
    def sample_presentation_data(self, sample_slide_info):
        """Create sample presentation data."""
        return {
            'slides': [
                {
                    'slide_number': 1,
                    'slide_info': sample_slide_info,
                    'object_counts': {'shapes': 2, 'text_boxes': 1, 'images': 0, 'tables': 0, 'charts': 0},
                    'notes': 'Sample notes',
                    'xml_content': '<xml>slide1</xml>'
                },
                {
                    'slide_number': 2,
                    'slide_info': sample_slide_info,
                    'object_counts': {'shapes': 1, 'text_boxes': 0, 'images': 1, 'tables': 0, 'charts': 0},
                    'notes': '',
                    'xml_content': '<xml>slide2</xml>'
                }
            ],
            'metadata': {'slide_count': 2},
            'sections': [{'name': 'Introduction', 'id': '1'}]
        }
    
    def test_estimate_presentation_duration(self, presentation_analyzer, sample_presentation_data):
        """Test presentation duration estimation."""
        duration = presentation_analyzer._estimate_presentation_duration(sample_presentation_data)
        
        assert 'estimated_minutes' in duration
        assert 'estimated_range' in duration
        assert 'slides_per_minute' in duration
        assert duration['estimated_minutes'] > 0
        assert isinstance(duration['slides_per_minute'], (int, float))
    
    def test_calculate_complexity_score(self, presentation_analyzer, sample_presentation_data):
        """Test complexity score calculation."""
        score = presentation_analyzer._calculate_complexity_score(sample_presentation_data)
        
        assert isinstance(score, float)
        assert 0.0 <= score <= 10.0
    
    def test_determine_slide_type_title_slide(self, presentation_analyzer):
        """Test slide type determination for title slide."""
        slide_info = Mock()
        slide_info.title = "Presentation Title"
        slide_info.subtitle = None
        slide_info.text_elements = []
        
        object_counts = {'shapes': 1, 'text_boxes': 1, 'images': 0, 'tables': 0, 'charts': 0}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.TITLE_SLIDE
        assert confidence > 0.8
        assert "minimal_content" in characteristics
        assert "title_only" in characteristics
    
    def test_determine_slide_type_section_header(self, presentation_analyzer):
        """Test slide type determination for section header."""
        slide_info = Mock()
        slide_info.title = "Section Title"
        slide_info.subtitle = "Section Subtitle"
        slide_info.text_elements = []
        
        object_counts = {'shapes': 2, 'text_boxes': 2, 'images': 0, 'tables': 0, 'charts': 0}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.SECTION_HEADER
        assert confidence > 0.7
        assert "title_and_subtitle" in characteristics
        assert "section_divider" in characteristics
    
    def test_determine_slide_type_table_slide(self, presentation_analyzer):
        """Test slide type determination for table slide."""
        slide_info = Mock()
        slide_info.title = "Data Table"
        slide_info.subtitle = None
        slide_info.text_elements = []
        
        object_counts = {'shapes': 1, 'text_boxes': 1, 'images': 0, 'tables': 2, 'charts': 0}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.TABLE_SLIDE
        assert confidence > 0.8
        assert "2_tables" in characteristics
        assert "table_dominant" in characteristics
    
    def test_determine_slide_type_chart_slide(self, presentation_analyzer):
        """Test slide type determination for chart slide."""
        slide_info = Mock()
        slide_info.title = "Sales Chart"
        slide_info.subtitle = None
        slide_info.text_elements = []
        
        object_counts = {'shapes': 1, 'text_boxes': 1, 'images': 0, 'tables': 0, 'charts': 2}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.CHART_SLIDE
        assert confidence > 0.8
        assert "2_charts" in characteristics
        assert "chart_dominant" in characteristics
    
    def test_determine_slide_type_image_slide(self, presentation_analyzer):
        """Test slide type determination for image slide."""
        slide_info = Mock()
        slide_info.title = "Product Images"
        slide_info.subtitle = None
        slide_info.text_elements = []
        
        object_counts = {'shapes': 1, 'text_boxes': 1, 'images': 3, 'tables': 0, 'charts': 0}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.IMAGE_SLIDE
        assert confidence > 0.7
        assert "3_images" in characteristics
        assert "image_dominant" in characteristics
    
    def test_determine_slide_type_bullet_slide(self, presentation_analyzer):
        """Test slide type determination for bullet slide."""
        slide_info = Mock()
        slide_info.title = "Key Points"
        slide_info.subtitle = None
        slide_info.text_elements = [
            {'content_plain': '• First point'},
            {'content_plain': '• Second point'},
            {'content_plain': '• Third point'}
        ]
        
        object_counts = {'shapes': 1, 'text_boxes': 3, 'images': 0, 'tables': 0, 'charts': 0}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.BULLET_SLIDE
        assert confidence > 0.6
        assert "bullet_points" in characteristics
        assert "3_text_elements" in characteristics
    
    def test_determine_slide_type_blank_slide(self, presentation_analyzer):
        """Test slide type determination for blank slide."""
        slide_info = Mock()
        slide_info.title = None
        slide_info.subtitle = None
        slide_info.text_elements = []
        
        object_counts = {'shapes': 0, 'text_boxes': 0, 'images': 0, 'tables': 0, 'charts': 0}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.BLANK_SLIDE
        assert confidence == 1.0
        assert "no_content" in characteristics
    
    def test_determine_slide_type_mixed_content(self, presentation_analyzer):
        """Test slide type determination for mixed content slide."""
        slide_info = Mock()
        slide_info.title = "Mixed Content"
        slide_info.subtitle = None
        slide_info.text_elements = [
            {'content_plain': 'Some text content'}
        ]
        
        object_counts = {'shapes': 2, 'text_boxes': 1, 'images': 1, 'tables': 1, 'charts': 1}
        
        slide_type, confidence, characteristics = presentation_analyzer._determine_slide_type(
            slide_info, object_counts
        )
        
        assert slide_type == SlideType.MIXED_CONTENT
        assert confidence > 0.5
        assert "mixed_content" in characteristics
        assert "5_content_types" in characteristics
    
    def test_generate_slide_content_summary(self, presentation_analyzer):
        """Test slide content summary generation."""
        slide_info = Mock()
        slide_info.title = "Test Slide Title"
        slide_info.text_elements = [
            {'content_plain': 'This is some sample text content for testing purposes.'}
        ]
        
        object_counts = {'shapes': 2, 'text_boxes': 1, 'images': 1, 'tables': 0, 'charts': 0}
        
        summary = presentation_analyzer._generate_slide_content_summary(slide_info, object_counts)
        
        assert "Title: Test Slide Title" in summary
        assert "Objects:" in summary
        assert "2 shapes" in summary
        assert "1 text_boxes" in summary
        assert "1 images" in summary
        assert "Text:" in summary
    
    def test_generate_slide_content_summary_empty(self, presentation_analyzer):
        """Test slide content summary generation for empty slide."""
        slide_info = Mock()
        slide_info.title = None
        slide_info.text_elements = []
        
        object_counts = {'shapes': 0, 'text_boxes': 0, 'images': 0, 'tables': 0, 'charts': 0}
        
        summary = presentation_analyzer._generate_slide_content_summary(slide_info, object_counts)
        
        assert summary == "Empty slide"
    
    def test_classify_single_slide(self, presentation_analyzer, sample_slide_info):
        """Test single slide classification."""
        slide_data = {
            'slide_number': 1,
            'slide_info': sample_slide_info,
            'object_counts': {'shapes': 2, 'text_boxes': 1, 'images': 0, 'tables': 0, 'charts': 0}
        }
        
        classification = presentation_analyzer._classify_single_slide(slide_data, AnalysisDepth.BASIC)
        
        assert isinstance(classification, SlideClassification)
        assert classification.slide_number == 1
        assert isinstance(classification.slide_type, SlideType)
        assert 0.0 <= classification.confidence <= 1.0
        assert isinstance(classification.characteristics, list)
        assert isinstance(classification.content_summary, str)
        assert isinstance(classification.object_counts, dict)
    
    def test_detect_structural_issues(self, presentation_analyzer):
        """Test structural issue detection."""
        # Create slide classifications without title slide
        slide_classifications = [
            SlideClassification(
                slide_number=1,
                slide_type=SlideType.CONTENT_SLIDE,
                confidence=0.8,
                characteristics=[],
                content_summary="Content",
                object_counts={}
            ),
            SlideClassification(
                slide_number=2,
                slide_type=SlideType.BLANK_SLIDE,
                confidence=1.0,
                characteristics=[],
                content_summary="Empty",
                object_counts={}
            )
        ]
        
        sections = []
        
        issues = presentation_analyzer._detect_structural_issues(slide_classifications, sections)
        
        assert "No title slide detected" in issues
        assert any("Blank slides found" in issue for issue in issues)
        assert any("Very short presentation" in issue for issue in issues)
    
    def test_calculate_readability_score(self, presentation_analyzer, sample_presentation_data):
        """Test readability score calculation."""
        score = presentation_analyzer._calculate_readability_score(sample_presentation_data)
        
        assert isinstance(score, float)
        assert 0.0 <= score <= 10.0
    
    def test_assess_content_density(self, presentation_analyzer):
        """Test content density assessment."""
        # Low density slides
        low_density_classifications = [
            SlideClassification(
                slide_number=1,
                slide_type=SlideType.CONTENT_SLIDE,
                confidence=0.8,
                characteristics=[],
                content_summary="",
                object_counts={'shapes': 1, 'text_boxes': 1}
            )
        ]
        
        density = presentation_analyzer._assess_content_density(low_density_classifications)
        assert density == "low"
        
        # High density slides
        high_density_classifications = [
            SlideClassification(
                slide_number=1,
                slide_type=SlideType.MIXED_CONTENT,
                confidence=0.8,
                characteristics=[],
                content_summary="",
                object_counts={'shapes': 5, 'text_boxes': 3, 'images': 2, 'tables': 1}
            )
        ]
        
        density = presentation_analyzer._assess_content_density(high_density_classifications)
        assert density == "high"
    
    def test_assess_visual_balance(self, presentation_analyzer):
        """Test visual balance assessment."""
        # Well balanced slides
        balanced_classifications = [
            SlideClassification(1, SlideType.TITLE_SLIDE, 0.9, [], "", {}),
            SlideClassification(2, SlideType.CONTENT_SLIDE, 0.8, [], "", {}),
            SlideClassification(3, SlideType.TABLE_SLIDE, 0.9, [], "", {}),
            SlideClassification(4, SlideType.CHART_SLIDE, 0.9, [], "", {}),
            SlideClassification(5, SlideType.IMAGE_SLIDE, 0.8, [], "", {})
        ]
        
        balance = presentation_analyzer._assess_visual_balance(balanced_classifications)
        assert balance in ["well_balanced", "balanced"]
        
        # Monotonous slides
        monotonous_classifications = [
            SlideClassification(1, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(2, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(3, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(4, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(5, SlideType.BULLET_SLIDE, 0.7, [], "", {})
        ]
        
        balance = presentation_analyzer._assess_visual_balance(monotonous_classifications)
        assert balance == "monotonous"
    
    def test_find_consistency_issues(self, presentation_analyzer, sample_presentation_data):
        """Test consistency issue detection."""
        slide_classifications = [
            SlideClassification(1, SlideType.CONTENT_SLIDE, 0.8, [], "", {'shapes': 2}),
            SlideClassification(2, SlideType.CONTENT_SLIDE, 0.8, [], "", {'shapes': 15})  # High variation
        ]
        
        issues = presentation_analyzer._find_consistency_issues(sample_presentation_data, slide_classifications)
        
        assert isinstance(issues, list)
        # Should detect large variation in content density
        assert any("variation in content density" in issue for issue in issues)
    
    def test_detect_title_patterns(self, presentation_analyzer):
        """Test title pattern detection."""
        slides = [
            {'slide_number': 1, 'slide_info': Mock(title="1. Introduction")},
            {'slide_number': 2, 'slide_info': Mock(title="2. Main Content")},
            {'slide_number': 3, 'slide_info': Mock(title="3. Conclusion")},
            {'slide_number': 4, 'slide_info': Mock(title="What is our goal?")},
            {'slide_number': 5, 'slide_info': Mock(title="How do we achieve it?")}
        ]
        
        patterns = presentation_analyzer._detect_title_patterns(slides)
        
        # Should detect numbered titles pattern
        numbered_pattern = next((p for p in patterns if p.pattern_name == "numbered_titles"), None)
        assert numbered_pattern is not None
        assert numbered_pattern.occurrences == 3
        
        # Should detect question titles pattern
        question_pattern = next((p for p in patterns if p.pattern_name == "question_titles"), None)
        assert question_pattern is not None
        assert question_pattern.occurrences == 2
    
    def test_generate_recommendations(self, presentation_analyzer, sample_presentation_data):
        """Test recommendation generation."""
        slide_classifications = [
            SlideClassification(1, SlideType.CONTENT_SLIDE, 0.8, [], "", {'shapes': 15}),  # High density
            SlideClassification(2, SlideType.BULLET_SLIDE, 0.7, [], "", {'shapes': 2}),
            SlideClassification(3, SlideType.BULLET_SLIDE, 0.7, [], "", {'shapes': 2})
        ]
        
        structure = PresentationStructure(
            total_slides=3,
            slide_types={'content_slide': 1, 'bullet_slide': 2},
            sections=[],
            content_flow=[],
            structural_issues=[]
        )
        
        consistency_issues = ["Test issue"]
        
        recommendations = presentation_analyzer._generate_recommendations(
            sample_presentation_data, slide_classifications, structure, consistency_issues
        )
        
        assert isinstance(recommendations, list)
        assert any("title slide" in rec for rec in recommendations)
        assert any("consistency" in rec for rec in recommendations)
    
    def test_identify_strengths(self, presentation_analyzer):
        """Test strength identification."""
        slide_classifications = [
            SlideClassification(1, SlideType.TITLE_SLIDE, 0.9, [], "", {}),
            SlideClassification(2, SlideType.CONTENT_SLIDE, 0.8, [], "", {}),
            SlideClassification(3, SlideType.TABLE_SLIDE, 0.9, [], "", {}),
            SlideClassification(4, SlideType.CHART_SLIDE, 0.9, [], "", {}),
            SlideClassification(5, SlideType.IMAGE_SLIDE, 0.8, [], "", {})
        ]
        
        structure = PresentationStructure(
            total_slides=5,
            slide_types={},
            sections=[],
            content_flow=[],
            structural_issues=[]
        )
        
        strengths = presentation_analyzer._identify_strengths(slide_classifications, structure)
        
        assert isinstance(strengths, list)
        assert any("Appropriate presentation length" in strength for strength in strengths)
        assert any("Good variety in slide types" in strength for strength in strengths)
        assert any("Good use of visual elements" in strength for strength in strengths)
    
    def test_identify_improvement_areas(self, presentation_analyzer):
        """Test improvement area identification."""
        # Create mostly bullet slides
        slide_classifications = [
            SlideClassification(1, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(2, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(3, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(4, SlideType.BULLET_SLIDE, 0.7, [], "", {}),
            SlideClassification(5, SlideType.CONTENT_SLIDE, 0.8, [], "", {})
        ]
        
        structure = PresentationStructure(
            total_slides=5,
            slide_types={},
            sections=[],
            content_flow=[],
            structural_issues=["Test issue"]
        )
        
        consistency_issues = ["Consistency issue"]
        
        areas = presentation_analyzer._identify_improvement_areas(
            slide_classifications, structure, consistency_issues
        )
        
        assert isinstance(areas, list)
        assert any("text-heavy slides" in area for area in areas)
        assert any("visual elements" in area for area in areas)
        assert any("structural issues" in area for area in areas)
        assert any("consistency" in area for area in areas)
    
    def test_collect_sample_content(self, presentation_analyzer, sample_presentation_data):
        """Test sample content collection."""
        sample_content = presentation_analyzer._collect_sample_content(
            sample_presentation_data, AnalysisDepth.DETAILED
        )
        
        assert isinstance(sample_content, dict)
        assert 'sample_titles' in sample_content
        assert 'sample_text' in sample_content
        assert isinstance(sample_content['sample_titles'], list)
        assert isinstance(sample_content['sample_text'], list)
    
    def test_cache_operations(self, presentation_analyzer):
        """Test cache operations."""
        # Add something to cache
        presentation_analyzer._analysis_cache["test_key"] = "test_value"
        assert len(presentation_analyzer._analysis_cache) == 1
        
        # Clear cache
        presentation_analyzer.clear_cache()
        assert len(presentation_analyzer._analysis_cache) == 0
    
    @patch('powerpoint_mcp_server.utils.zip_extractor.ZipExtractor')
    @pytest.mark.asyncio
    async def test_analyze_presentation_integration(self, mock_zip_extractor, presentation_analyzer, sample_slide_info):
        """Test the main analyze_presentation method integration."""
        # Mock ZipExtractor
        mock_extractor_instance = Mock()
        mock_zip_extractor.return_value.__enter__.return_value = mock_extractor_instance
        
        mock_extractor_instance.get_slide_xml_files.return_value = ["slide1.xml"]
        mock_extractor_instance.read_xml_content.side_effect = [
            "<xml>presentation</xml>",  # presentation.xml
            "<xml>slide1</xml>",        # slide1.xml
            None                        # notes (not found)
        ]
        
        # Mock content extractor methods
        presentation_analyzer.content_extractor.extract_slide_content.return_value = sample_slide_info
        
        # Test analysis
        result = await presentation_analyzer.analyze_presentation(
            file_path="test.pptx",
            analysis_depth=AnalysisDepth.BASIC,
            include_sample_content=True
        )
        
        assert isinstance(result, PresentationOverview)
        assert result.file_path == "test.pptx"
        assert result.analysis_depth == AnalysisDepth.BASIC
        assert isinstance(result.metadata, dict)
        assert isinstance(result.structure, PresentationStructure)
        assert isinstance(result.slide_classifications, list)
        assert isinstance(result.content_patterns, list)
        assert isinstance(result.insights, PresentationInsights)
        assert isinstance(result.sample_content, dict)


class TestSlideClassification:
    """Test cases for SlideClassification class."""
    
    def test_slide_classification_creation(self):
        """Test creating a SlideClassification."""
        classification = SlideClassification(
            slide_number=1,
            slide_type=SlideType.TITLE_SLIDE,
            confidence=0.9,
            characteristics=["minimal_content", "title_only"],
            content_summary="Title: Welcome",
            object_counts={'shapes': 1, 'text_boxes': 1}
        )
        
        assert classification.slide_number == 1
        assert classification.slide_type == SlideType.TITLE_SLIDE
        assert classification.confidence == 0.9
        assert classification.characteristics == ["minimal_content", "title_only"]
        assert classification.content_summary == "Title: Welcome"
        assert classification.object_counts == {'shapes': 1, 'text_boxes': 1}


class TestContentPattern:
    """Test cases for ContentPattern class."""
    
    def test_content_pattern_creation(self):
        """Test creating a ContentPattern."""
        pattern = ContentPattern(
            pattern_type="title_structure",
            pattern_name="numbered_titles",
            occurrences=5,
            slides=[1, 2, 3, 4, 5],
            examples=["1. Introduction", "2. Main Content"],
            confidence=0.9
        )
        
        assert pattern.pattern_type == "title_structure"
        assert pattern.pattern_name == "numbered_titles"
        assert pattern.occurrences == 5
        assert pattern.slides == [1, 2, 3, 4, 5]
        assert pattern.examples == ["1. Introduction", "2. Main Content"]
        assert pattern.confidence == 0.9


class TestPresentationStructure:
    """Test cases for PresentationStructure class."""
    
    def test_presentation_structure_creation(self):
        """Test creating a PresentationStructure."""
        structure = PresentationStructure(
            total_slides=10,
            slide_types={'title_slide': 1, 'content_slide': 8, 'conclusion_slide': 1},
            sections=[{'name': 'Introduction', 'id': '1'}],
            content_flow=['title_slide', 'content_slide', 'content_slide'],
            structural_issues=['No conclusion slide']
        )
        
        assert structure.total_slides == 10
        assert structure.slide_types == {'title_slide': 1, 'content_slide': 8, 'conclusion_slide': 1}
        assert structure.sections == [{'name': 'Introduction', 'id': '1'}]
        assert structure.content_flow == ['title_slide', 'content_slide', 'content_slide']
        assert structure.structural_issues == ['No conclusion slide']


class TestPresentationInsights:
    """Test cases for PresentationInsights class."""
    
    def test_presentation_insights_creation(self):
        """Test creating PresentationInsights."""
        insights = PresentationInsights(
            readability_score=7.5,
            content_density="medium",
            visual_balance="balanced",
            consistency_issues=["Inconsistent title usage"],
            recommendations=["Add more visual elements"],
            strengths=["Good structure"],
            areas_for_improvement=["Reduce text density"]
        )
        
        assert insights.readability_score == 7.5
        assert insights.content_density == "medium"
        assert insights.visual_balance == "balanced"
        assert insights.consistency_issues == ["Inconsistent title usage"]
        assert insights.recommendations == ["Add more visual elements"]
        assert insights.strengths == ["Good structure"]
        assert insights.areas_for_improvement == ["Reduce text density"]


if __name__ == "__main__":
    pytest.main([__file__])