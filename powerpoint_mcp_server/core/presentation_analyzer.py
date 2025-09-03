"""
Presentation overview and analysis system for comprehensive presentation insights.
"""

import re
import logging
from typing import Dict, List, Any, Optional, Union, Tuple
from dataclasses import dataclass, field
from enum import Enum
from collections import defaultdict, Counter

from .content_extractor import ContentExtractor
from ..utils.zip_extractor import ZipExtractor

logger = logging.getLogger(__name__)


class AnalysisDepth(Enum):
    """Enumeration of analysis depth levels."""
    BASIC = "basic"
    DETAILED = "detailed"
    COMPREHENSIVE = "comprehensive"


class SlideType(Enum):
    """Enumeration of slide types."""
    TITLE_SLIDE = "title_slide"
    CONTENT_SLIDE = "content_slide"
    SECTION_HEADER = "section_header"
    BULLET_SLIDE = "bullet_slide"
    TABLE_SLIDE = "table_slide"
    IMAGE_SLIDE = "image_slide"
    CHART_SLIDE = "chart_slide"
    MIXED_CONTENT = "mixed_content"
    BLANK_SLIDE = "blank_slide"
    UNKNOWN = "unknown"


@dataclass
class SlideClassification:
    """Classification of a single slide."""
    slide_number: int
    slide_type: SlideType
    confidence: float
    characteristics: List[str]
    content_summary: str
    object_counts: Dict[str, int]


@dataclass
class ContentPattern:
    """A detected content pattern in the presentation."""
    pattern_type: str
    pattern_name: str
    occurrences: int
    slides: List[int]
    examples: List[str]
    confidence: float


@dataclass
class PresentationStructure:
    """Structure analysis of the presentation."""
    total_slides: int
    slide_types: Dict[str, int]
    sections: List[Dict[str, Any]]
    content_flow: List[str]
    structural_issues: List[str]


@dataclass
class PresentationInsights:
    """Insights and recommendations for the presentation."""
    readability_score: float
    content_density: str
    visual_balance: str
    consistency_issues: List[str]
    recommendations: List[str]
    strengths: List[str]
    areas_for_improvement: List[str]


@dataclass
class PresentationOverview:
    """Complete presentation overview and analysis."""
    file_path: str
    metadata: Dict[str, Any]
    structure: PresentationStructure
    slide_classifications: List[SlideClassification]
    content_patterns: List[ContentPattern]
    insights: PresentationInsights
    analysis_depth: AnalysisDepth
    sample_content: Dict[str, Any] = field(default_factory=dict)


class PresentationAnalyzer:
    """
    Analyzer for comprehensive presentation overview and analysis.
    """
    
    def __init__(self, content_extractor: Optional[ContentExtractor] = None):
        """Initialize the presentation analyzer."""
        self.content_extractor = content_extractor or ContentExtractor()
        self._analysis_cache = {}
    
    async def analyze_presentation(
        self,
        file_path: str,
        analysis_depth: AnalysisDepth = AnalysisDepth.BASIC,
        include_sample_content: bool = True
    ) -> PresentationOverview:
        """
        Analyze a presentation for comprehensive overview and insights.
        
        Args:
            file_path: Path to the PowerPoint file
            analysis_depth: Depth of analysis to perform
            include_sample_content: Whether to include sample content
            
        Returns:
            PresentationOverview with analysis results
        """
        logger.info(f"Analyzing presentation {file_path} with depth {analysis_depth.value}")
        
        try:
            # Extract presentation data
            presentation_data = self._extract_presentation_data(file_path)
            
            # Analyze presentation metadata
            metadata = self._analyze_metadata(presentation_data)
            
            # Classify slides
            slide_classifications = self._classify_slides(presentation_data, analysis_depth)
            
            # Analyze structure
            structure = self._analyze_structure(presentation_data, slide_classifications)
            
            # Detect content patterns
            content_patterns = []
            if analysis_depth in [AnalysisDepth.DETAILED, AnalysisDepth.COMPREHENSIVE]:
                content_patterns = self._detect_content_patterns(presentation_data, slide_classifications)
            
            # Generate insights
            insights = PresentationInsights(
                readability_score=0.0,
                content_density="medium",
                visual_balance="balanced",
                consistency_issues=[],
                recommendations=[],
                strengths=[],
                areas_for_improvement=[]
            )
            
            if analysis_depth == AnalysisDepth.COMPREHENSIVE:
                insights = self._generate_insights(presentation_data, slide_classifications, structure)
            
            # Collect sample content
            sample_content = {}
            if include_sample_content:
                sample_content = self._collect_sample_content(presentation_data, analysis_depth)
            
            overview = PresentationOverview(
                file_path=file_path,
                metadata=metadata,
                structure=structure,
                slide_classifications=slide_classifications,
                content_patterns=content_patterns,
                insights=insights,
                analysis_depth=analysis_depth,
                sample_content=sample_content
            )
            
            logger.info(f"Analysis complete: {len(slide_classifications)} slides analyzed")
            return overview
            
        except Exception as e:
            logger.error(f"Error analyzing presentation: {e}")
            raise
    
    def _extract_presentation_data(self, file_path: str) -> Dict[str, Any]:
        """Extract all necessary data from the presentation."""
        try:
            presentation_data = {
                'slides': [],
                'metadata': {},
                'sections': []
            }
            
            with ZipExtractor(file_path) as extractor:
                # Get presentation metadata
                presentation_xml = extractor.read_xml_content('ppt/presentation.xml')
                if presentation_xml:
                    presentation_data['metadata'] = self.content_extractor.extract_presentation_metadata(presentation_xml)
                    presentation_data['sections'] = self.content_extractor.extract_section_information(presentation_xml)
                
                # Get slide data
                slide_files = extractor.get_slide_xml_files()
                
                for i, slide_file in enumerate(slide_files, 1):
                    slide_xml = extractor.read_xml_content(slide_file)
                    if slide_xml:
                        # Extract comprehensive slide content
                        slide_info = self.content_extractor.extract_slide_content(slide_xml, i)
                        
                        # Get object counts
                        root = self.content_extractor.xml_parser.parse_xml_string(slide_xml)
                        object_counts = self.content_extractor._count_slide_objects(root) if root else {}
                        
                        # Get notes if available
                        notes_file = f'ppt/notesSlides/notesSlide{i}.xml'
                        notes_content = ""
                        try:
                            notes_xml = extractor.read_xml_content(notes_file)
                            if notes_xml:
                                notes_content = self.content_extractor.extract_slide_notes(notes_xml)
                        except Exception as e:
                            logger.debug(f"Notes not available for slide {i}: {e}")
                            # Notes are optional, so we continue without them
                        
                        slide_data = {
                            'slide_number': i,
                            'slide_info': slide_info,
                            'object_counts': object_counts,
                            'notes': notes_content,
                            'xml_content': slide_xml
                        }
                        
                        presentation_data['slides'].append(slide_data)
            
            return presentation_data
            
        except Exception as e:
            logger.warning(f"Failed to extract presentation data: {e}")
            return {'slides': [], 'metadata': {}, 'sections': []}
    
    def _analyze_metadata(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze presentation metadata."""
        try:
            metadata = presentation_data.get('metadata', {})
            
            # Calculate notes statistics
            slides = presentation_data.get('slides', [])
            notes_stats = {
                'slides_with_notes': sum(1 for slide in slides if slide.get('notes', '').strip()),
                'total_notes_length': sum(len(slide.get('notes', '')) for slide in slides),
                'average_notes_length': 0
            }
            if notes_stats['slides_with_notes'] > 0:
                notes_stats['average_notes_length'] = notes_stats['total_notes_length'] / notes_stats['slides_with_notes']
            
            analysis = {
                'slide_count': len(slides),
                'slide_size': metadata.get('slide_size'),
                'has_sections': len(presentation_data.get('sections', [])) > 0,
                'section_count': len(presentation_data.get('sections', [])),
                'sections': presentation_data.get('sections', []),
                'notes_statistics': notes_stats,
                'estimated_duration': self._estimate_presentation_duration(presentation_data),
                'complexity_score': self._calculate_complexity_score(presentation_data)
            }
            
            # Add original metadata
            analysis.update(metadata)
            
            return analysis
            
        except Exception as e:
            logger.warning(f"Failed to analyze metadata: {e}")
            return {}
    
    def _estimate_presentation_duration(self, presentation_data: Dict[str, Any]) -> Dict[str, Any]:
        """Estimate presentation duration based on content."""
        try:
            slides = presentation_data.get('slides', [])
            
            # Basic estimation: 1-2 minutes per slide depending on content
            total_minutes = 0
            
            for slide_data in slides:
                slide_info = slide_data.get('slide_info')
                object_counts = slide_data.get('object_counts', {})
                
                # Base time per slide
                slide_minutes = 1.0
                
                # Add time for tables
                if object_counts.get('tables', 0) > 0:
                    slide_minutes += 0.5 * object_counts['tables']
                
                # Add time for charts
                if object_counts.get('charts', 0) > 0:
                    slide_minutes += 0.5 * object_counts['charts']
                
                # Add time for text content
                text_elements = slide_info.text_elements if slide_info else []
                total_text_length = sum(len(elem.get('content_plain', '')) for elem in text_elements)
                if total_text_length > 200:
                    slide_minutes += 0.5
                
                # Add time for speaker notes
                notes = slide_data.get('notes', '')
                if notes and len(notes) > 100:
                    slide_minutes += 0.3
                
                total_minutes += slide_minutes
            
            return {
                'estimated_minutes': round(total_minutes, 1),
                'estimated_range': f"{round(total_minutes * 0.8, 1)}-{round(total_minutes * 1.2, 1)} minutes",
                'slides_per_minute': round(len(slides) / total_minutes, 2) if total_minutes > 0 else 0
            }
            
        except Exception as e:
            logger.warning(f"Failed to estimate duration: {e}")
            return {'estimated_minutes': 0}
    
    def _calculate_complexity_score(self, presentation_data: Dict[str, Any]) -> float:
        """Calculate a complexity score for the presentation."""
        try:
            slides = presentation_data.get('slides', [])
            if not slides:
                return 0.0
            
            complexity_factors = []
            
            for slide_data in slides:
                object_counts = slide_data.get('object_counts', {})
                slide_info = slide_data.get('slide_info')
                
                # Factor 1: Number of objects
                total_objects = sum(object_counts.values())
                object_complexity = min(total_objects / 5.0, 1.0)  # Normalize to 0-1
                
                # Factor 2: Text complexity
                text_complexity = 0.0
                if slide_info and slide_info.text_elements:
                    total_text = sum(len(elem.get('content_plain', '')) for elem in slide_info.text_elements)
                    text_complexity = min(total_text / 500.0, 1.0)  # Normalize to 0-1
                
                # Factor 3: Table complexity
                table_complexity = min(object_counts.get('tables', 0) / 2.0, 1.0)
                
                # Factor 4: Chart complexity
                chart_complexity = min(object_counts.get('charts', 0) / 2.0, 1.0)
                
                slide_complexity = (object_complexity + text_complexity + table_complexity + chart_complexity) / 4.0
                complexity_factors.append(slide_complexity)
            
            # Average complexity across all slides
            avg_complexity = sum(complexity_factors) / len(complexity_factors)
            
            # Scale to 0-10
            return round(avg_complexity * 10, 1)
            
        except Exception as e:
            logger.warning(f"Failed to calculate complexity score: {e}")
            return 0.0
    
    def _classify_slides(
        self,
        presentation_data: Dict[str, Any],
        analysis_depth: AnalysisDepth
    ) -> List[SlideClassification]:
        """Classify each slide by type and characteristics."""
        try:
            slides = presentation_data.get('slides', [])
            classifications = []
            
            for slide_data in slides:
                classification = self._classify_single_slide(slide_data, analysis_depth)
                classifications.append(classification)
            
            return classifications
            
        except Exception as e:
            logger.warning(f"Failed to classify slides: {e}")
            return []
    
    def _classify_single_slide(self, slide_data: Dict[str, Any], analysis_depth: AnalysisDepth) -> SlideClassification:
        """Classify a single slide."""
        try:
            slide_number = slide_data.get('slide_number', 0)
            slide_info = slide_data.get('slide_info')
            object_counts = slide_data.get('object_counts', {})
            
            # Determine slide type
            slide_type, confidence, characteristics = self._determine_slide_type(slide_info, object_counts)
            
            # Generate content summary
            content_summary = self._generate_slide_content_summary(slide_info, object_counts)
            
            return SlideClassification(
                slide_number=slide_number,
                slide_type=slide_type,
                confidence=confidence,
                characteristics=characteristics,
                content_summary=content_summary,
                object_counts=object_counts
            )
            
        except Exception as e:
            logger.warning(f"Failed to classify slide: {e}")
            return SlideClassification(
                slide_number=slide_data.get('slide_number', 0),
                slide_type=SlideType.UNKNOWN,
                confidence=0.0,
                characteristics=[],
                content_summary="",
                object_counts={}
            )
    
    def _determine_slide_type(
        self,
        slide_info,
        object_counts: Dict[str, int]
    ) -> Tuple[SlideType, float, List[str]]:
        """Determine the type of a slide based on its content."""
        try:
            characteristics = []
            
            # Check for title slide
            if slide_info and slide_info.title and not slide_info.subtitle:
                if object_counts.get('shapes', 0) <= 2 and sum(object_counts.values()) <= 3:
                    characteristics.append("minimal_content")
                    characteristics.append("title_only")
                    return SlideType.TITLE_SLIDE, 0.9, characteristics
            
            # Check for section header
            if slide_info and slide_info.title and slide_info.subtitle:
                if sum(object_counts.values()) <= 4:
                    characteristics.append("title_and_subtitle")
                    characteristics.append("section_divider")
                    return SlideType.SECTION_HEADER, 0.8, characteristics
            
            # Check for table slide
            if object_counts.get('tables', 0) > 0:
                characteristics.append(f"{object_counts['tables']}_tables")
                if object_counts['tables'] >= sum(object_counts.values()) * 0.5:
                    characteristics.append("table_dominant")
                    return SlideType.TABLE_SLIDE, 0.9, characteristics
            
            # Check for chart slide
            if object_counts.get('charts', 0) > 0:
                characteristics.append(f"{object_counts['charts']}_charts")
                if object_counts['charts'] >= sum(object_counts.values()) * 0.4:
                    characteristics.append("chart_dominant")
                    return SlideType.CHART_SLIDE, 0.9, characteristics
            
            # Check for image slide
            if object_counts.get('images', 0) > 0:
                characteristics.append(f"{object_counts['images']}_images")
                if object_counts['images'] >= sum(object_counts.values()) * 0.4:
                    characteristics.append("image_dominant")
                    return SlideType.IMAGE_SLIDE, 0.8, characteristics
            
            # Check for bullet slide
            if slide_info and slide_info.text_elements:
                text_content = ' '.join(elem.get('content_plain', '') for elem in slide_info.text_elements)
                bullet_indicators = text_content.count('•') + text_content.count('-') + text_content.count('*')
                if bullet_indicators >= 3 or len(slide_info.text_elements) >= 3:
                    characteristics.append("bullet_points")
                    characteristics.append(f"{len(slide_info.text_elements)}_text_elements")
                    return SlideType.BULLET_SLIDE, 0.7, characteristics
            
            # Check for blank slide
            if sum(object_counts.values()) == 0:
                characteristics.append("no_content")
                return SlideType.BLANK_SLIDE, 1.0, characteristics
            
            # Check for mixed content
            content_types = sum(1 for count in object_counts.values() if count > 0)
            if content_types >= 3:
                characteristics.append("mixed_content")
                characteristics.append(f"{content_types}_content_types")
                return SlideType.MIXED_CONTENT, 0.6, characteristics
            
            # Default to content slide
            characteristics.append("standard_content")
            if slide_info and slide_info.title:
                characteristics.append("has_title")
            
            return SlideType.CONTENT_SLIDE, 0.5, characteristics
            
        except Exception as e:
            logger.warning(f"Failed to determine slide type: {e}")
            return SlideType.UNKNOWN, 0.0, []
    
    def _generate_slide_content_summary(self, slide_info, object_counts: Dict[str, int]) -> str:
        """Generate a content summary for a slide."""
        try:
            summary_parts = []
            
            # Add title if available
            if slide_info and slide_info.title:
                title_preview = slide_info.title[:50] + "..." if len(slide_info.title) > 50 else slide_info.title
                summary_parts.append(f"Title: {title_preview}")
            
            # Add object counts
            object_descriptions = []
            for obj_type, count in object_counts.items():
                if count > 0:
                    object_descriptions.append(f"{count} {obj_type}")
            
            if object_descriptions:
                summary_parts.append(f"Objects: {', '.join(object_descriptions)}")
            
            # Add text preview
            if slide_info and slide_info.text_elements:
                all_text = ' '.join(elem.get('content_plain', '') for elem in slide_info.text_elements)
                if all_text.strip():
                    text_preview = all_text[:100] + "..." if len(all_text) > 100 else all_text
                    summary_parts.append(f"Text: {text_preview}")
            
            return " | ".join(summary_parts) if summary_parts else "Empty slide"
            
        except Exception as e:
            logger.warning(f"Failed to generate content summary: {e}")
            return "Content summary unavailable"
    
    def _analyze_structure(
        self,
        presentation_data: Dict[str, Any],
        slide_classifications: List[SlideClassification]
    ) -> PresentationStructure:
        """Analyze the overall structure of the presentation."""
        try:
            slides = presentation_data.get('slides', [])
            sections = presentation_data.get('sections', [])
            
            # Count slide types
            slide_types = defaultdict(int)
            for classification in slide_classifications:
                slide_types[classification.slide_type.value] += 1
            
            # Analyze content flow
            content_flow = [classification.slide_type.value for classification in slide_classifications]
            
            # Detect structural issues
            structural_issues = self._detect_structural_issues(slide_classifications, sections)
            
            return PresentationStructure(
                total_slides=len(slides),
                slide_types=dict(slide_types),
                sections=[{'name': section.get('name', 'Unnamed'), 'id': section.get('id', '')} for section in sections],
                content_flow=content_flow,
                structural_issues=structural_issues
            )
            
        except Exception as e:
            logger.warning(f"Failed to analyze structure: {e}")
            return PresentationStructure(
                total_slides=0,
                slide_types={},
                sections=[],
                content_flow=[],
                structural_issues=[]
            )
    
    def _detect_structural_issues(
        self,
        slide_classifications: List[SlideClassification],
        sections: List[Dict[str, Any]]
    ) -> List[str]:
        """Detect potential structural issues in the presentation."""
        try:
            issues = []
            
            # Check for missing title slide
            if not any(cls.slide_type == SlideType.TITLE_SLIDE for cls in slide_classifications):
                issues.append("No title slide detected")
            
            # Check for too many consecutive slides of the same type
            consecutive_count = 1
            prev_type = None
            
            for classification in slide_classifications:
                if classification.slide_type == prev_type:
                    consecutive_count += 1
                else:
                    if consecutive_count > 5 and prev_type not in [SlideType.CONTENT_SLIDE]:
                        issues.append(f"Too many consecutive {prev_type.value} slides ({consecutive_count})")
                    consecutive_count = 1
                    prev_type = classification.slide_type
            
            # Check for blank slides
            blank_slides = [cls.slide_number for cls in slide_classifications if cls.slide_type == SlideType.BLANK_SLIDE]
            if blank_slides:
                issues.append(f"Blank slides found: {blank_slides}")
            
            # Check for very short presentation
            if len(slide_classifications) < 3:
                issues.append("Very short presentation (less than 3 slides)")
            
            # Check for very long presentation
            if len(slide_classifications) > 50:
                issues.append("Very long presentation (more than 50 slides)")
            
            return issues
            
        except Exception as e:
            logger.warning(f"Failed to detect structural issues: {e}")
            return []
    
    def _detect_content_patterns(
        self,
        presentation_data: Dict[str, Any],
        slide_classifications: List[SlideClassification]
    ) -> List[ContentPattern]:
        """Detect content patterns in the presentation."""
        try:
            patterns = []
            slides = presentation_data.get('slides', [])
            
            # Pattern 1: Title patterns
            title_patterns = self._detect_title_patterns(slides)
            patterns.extend(title_patterns)
            
            # Pattern 2: Layout patterns
            layout_patterns = self._detect_layout_patterns(slide_classifications)
            patterns.extend(layout_patterns)
            
            # Pattern 3: Content structure patterns
            structure_patterns = self._detect_structure_patterns(slides)
            patterns.extend(structure_patterns)
            
            return patterns
            
        except Exception as e:
            logger.warning(f"Failed to detect content patterns: {e}")
            return []
    
    def _detect_title_patterns(self, slides: List[Dict[str, Any]]) -> List[ContentPattern]:
        """Detect patterns in slide titles."""
        try:
            patterns = []
            titles = []
            
            for slide_data in slides:
                slide_info = slide_data.get('slide_info')
                if slide_info and slide_info.title:
                    titles.append((slide_data['slide_number'], slide_info.title))
            
            if not titles:
                return patterns
            
            # Pattern: Numbered titles
            numbered_titles = []
            for slide_num, title in titles:
                if re.match(r'^\d+\.', title.strip()):
                    numbered_titles.append((slide_num, title))
            
            if len(numbered_titles) >= 3:
                patterns.append(ContentPattern(
                    pattern_type="title_structure",
                    pattern_name="numbered_titles",
                    occurrences=len(numbered_titles),
                    slides=[slide_num for slide_num, _ in numbered_titles],
                    examples=[title[:50] for _, title in numbered_titles[:3]],
                    confidence=0.9
                ))
            
            # Pattern: Question titles
            question_titles = []
            for slide_num, title in titles:
                if title.strip().endswith('?'):
                    question_titles.append((slide_num, title))
            
            if len(question_titles) >= 2:
                patterns.append(ContentPattern(
                    pattern_type="title_structure",
                    pattern_name="question_titles",
                    occurrences=len(question_titles),
                    slides=[slide_num for slide_num, _ in question_titles],
                    examples=[title[:50] for _, title in question_titles[:3]],
                    confidence=0.8
                ))
            
            # Pattern: Common prefixes
            prefix_groups = defaultdict(list)
            for slide_num, title in titles:
                words = title.split()
                if len(words) >= 2:
                    prefix = ' '.join(words[:2])
                    prefix_groups[prefix].append((slide_num, title))
            
            for prefix, group in prefix_groups.items():
                if len(group) >= 3:
                    patterns.append(ContentPattern(
                        pattern_type="title_structure",
                        pattern_name="common_prefix",
                        occurrences=len(group),
                        slides=[slide_num for slide_num, _ in group],
                        examples=[f"Prefix: '{prefix}' - {title[:30]}" for _, title in group[:3]],
                        confidence=0.7
                    ))
            
            return patterns
            
        except Exception as e:
            logger.warning(f"Failed to detect title patterns: {e}")
            return []
    
    def _detect_layout_patterns(self, slide_classifications: List[SlideClassification]) -> List[ContentPattern]:
        """Detect patterns in slide layouts."""
        try:
            patterns = []
            
            # Count slide type sequences
            type_sequences = []
            for i in range(len(slide_classifications) - 1):
                current_type = slide_classifications[i].slide_type.value
                next_type = slide_classifications[i + 1].slide_type.value
                type_sequences.append(f"{current_type} -> {next_type}")
            
            # Find common sequences
            sequence_counts = Counter(type_sequences)
            for sequence, count in sequence_counts.items():
                if count >= 3:
                    patterns.append(ContentPattern(
                        pattern_type="layout_sequence",
                        pattern_name="repeated_sequence",
                        occurrences=count,
                        slides=[],  # Would need more complex tracking
                        examples=[sequence],
                        confidence=0.8
                    ))
            
            return patterns
            
        except Exception as e:
            logger.warning(f"Failed to detect layout patterns: {e}")
            return []
    
    def _detect_structure_patterns(self, slides: List[Dict[str, Any]]) -> List[ContentPattern]:
        """Detect patterns in content structure."""
        try:
            patterns = []
            
            # Pattern: Consistent bullet point usage
            bullet_slides = []
            for slide_data in slides:
                slide_info = slide_data.get('slide_info')
                if slide_info and slide_info.text_elements:
                    text_content = ' '.join(elem.get('content_plain', '') for elem in slide_info.text_elements)
                    bullet_count = text_content.count('•') + text_content.count('-') + text_content.count('*')
                    if bullet_count >= 3:
                        bullet_slides.append(slide_data['slide_number'])
            
            if len(bullet_slides) >= 5:
                patterns.append(ContentPattern(
                    pattern_type="content_structure",
                    pattern_name="consistent_bullets",
                    occurrences=len(bullet_slides),
                    slides=bullet_slides,
                    examples=["Consistent use of bullet points"],
                    confidence=0.8
                ))
            
            return patterns
            
        except Exception as e:
            logger.warning(f"Failed to detect structure patterns: {e}")
            return []
    
    def _generate_insights(
        self,
        presentation_data: Dict[str, Any],
        slide_classifications: List[SlideClassification],
        structure: PresentationStructure
    ) -> PresentationInsights:
        """Generate comprehensive insights and recommendations."""
        try:
            # Calculate readability score
            readability_score = self._calculate_readability_score(presentation_data)
            
            # Assess content density
            content_density = self._assess_content_density(slide_classifications)
            
            # Assess visual balance
            visual_balance = self._assess_visual_balance(slide_classifications)
            
            # Find consistency issues
            consistency_issues = self._find_consistency_issues(presentation_data, slide_classifications)
            
            # Generate recommendations
            recommendations = self._generate_recommendations(
                presentation_data, slide_classifications, structure, consistency_issues
            )
            
            # Identify strengths
            strengths = self._identify_strengths(slide_classifications, structure)
            
            # Identify areas for improvement
            areas_for_improvement = self._identify_improvement_areas(
                slide_classifications, structure, consistency_issues
            )
            
            return PresentationInsights(
                readability_score=readability_score,
                content_density=content_density,
                visual_balance=visual_balance,
                consistency_issues=consistency_issues,
                recommendations=recommendations,
                strengths=strengths,
                areas_for_improvement=areas_for_improvement
            )
            
        except Exception as e:
            logger.warning(f"Failed to generate insights: {e}")
            return PresentationInsights(
                readability_score=0.0,
                content_density="unknown",
                visual_balance="unknown",
                consistency_issues=[],
                recommendations=[],
                strengths=[],
                areas_for_improvement=[]
            )
    
    def _calculate_readability_score(self, presentation_data: Dict[str, Any]) -> float:
        """Calculate a readability score for the presentation."""
        try:
            slides = presentation_data.get('slides', [])
            if not slides:
                return 0.0
            
            readability_factors = []
            
            for slide_data in slides:
                slide_info = slide_data.get('slide_info')
                if not slide_info:
                    continue
                
                # Factor 1: Text length per slide
                total_text = 0
                if slide_info.text_elements:
                    total_text = sum(len(elem.get('content_plain', '')) for elem in slide_info.text_elements)
                
                text_score = 1.0 if total_text <= 200 else max(0.0, 1.0 - (total_text - 200) / 500)
                
                # Factor 2: Number of text elements
                text_elements_count = len(slide_info.text_elements) if slide_info.text_elements else 0
                elements_score = 1.0 if text_elements_count <= 5 else max(0.0, 1.0 - (text_elements_count - 5) / 10)
                
                # Factor 3: Title presence
                title_score = 1.0 if slide_info.title else 0.5
                
                slide_readability = (text_score + elements_score + title_score) / 3.0
                readability_factors.append(slide_readability)
            
            # Average readability across all slides
            avg_readability = sum(readability_factors) / len(readability_factors)
            
            # Scale to 0-10
            return round(avg_readability * 10, 1)
            
        except Exception as e:
            logger.warning(f"Failed to calculate readability score: {e}")
            return 0.0
    
    def _assess_content_density(self, slide_classifications: List[SlideClassification]) -> str:
        """Assess the content density of the presentation."""
        try:
            if not slide_classifications:
                return "unknown"
            
            # Calculate average objects per slide
            total_objects = sum(sum(cls.object_counts.values()) for cls in slide_classifications)
            avg_objects = total_objects / len(slide_classifications)
            
            if avg_objects <= 3:
                return "low"
            elif avg_objects <= 6:
                return "medium"
            else:
                return "high"
                
        except Exception as e:
            logger.warning(f"Failed to assess content density: {e}")
            return "unknown"
    
    def _assess_visual_balance(self, slide_classifications: List[SlideClassification]) -> str:
        """Assess the visual balance of the presentation."""
        try:
            if not slide_classifications:
                return "unknown"
            
            # Check for variety in slide types
            slide_types = set(cls.slide_type for cls in slide_classifications)
            type_variety = len(slide_types)
            
            # Check for extreme imbalances
            type_counts = Counter(cls.slide_type for cls in slide_classifications)
            max_type_ratio = max(type_counts.values()) / len(slide_classifications)
            
            if type_variety >= 4 and max_type_ratio <= 0.7:
                return "well_balanced"
            elif type_variety >= 3 and max_type_ratio <= 0.8:
                return "balanced"
            elif max_type_ratio >= 0.9:
                return "monotonous"
            else:
                return "somewhat_balanced"
                
        except Exception as e:
            logger.warning(f"Failed to assess visual balance: {e}")
            return "unknown"
    
    def _find_consistency_issues(
        self,
        presentation_data: Dict[str, Any],
        slide_classifications: List[SlideClassification]
    ) -> List[str]:
        """Find consistency issues in the presentation."""
        try:
            issues = []
            slides = presentation_data.get('slides', [])
            
            # Check title consistency
            titled_slides = []
            untitled_slides = []
            
            for slide_data in slides:
                slide_info = slide_data.get('slide_info')
                if slide_info and slide_info.title:
                    titled_slides.append(slide_data['slide_number'])
                else:
                    untitled_slides.append(slide_data['slide_number'])
            
            if len(untitled_slides) > 0 and len(titled_slides) > len(untitled_slides):
                issues.append(f"Inconsistent title usage: {len(untitled_slides)} slides without titles")
            
            # Check for extreme variations in content density
            object_counts = [sum(cls.object_counts.values()) for cls in slide_classifications]
            if object_counts:
                min_objects = min(object_counts)
                max_objects = max(object_counts)
                if max_objects > min_objects * 5 and max_objects > 10:
                    issues.append("Large variation in content density between slides")
            
            return issues
            
        except Exception as e:
            logger.warning(f"Failed to find consistency issues: {e}")
            return []
    
    def _generate_recommendations(
        self,
        presentation_data: Dict[str, Any],
        slide_classifications: List[SlideClassification],
        structure: PresentationStructure,
        consistency_issues: List[str]
    ) -> List[str]:
        """Generate actionable recommendations."""
        try:
            recommendations = []
            
            # Recommendations based on structure
            if not any(cls.slide_type == SlideType.TITLE_SLIDE for cls in slide_classifications):
                recommendations.append("Add a title slide to introduce your presentation")
            
            if structure.total_slides > 30:
                recommendations.append("Consider breaking this into multiple shorter presentations")
            
            # Recommendations based on content density
            high_density_slides = [
                cls.slide_number for cls in slide_classifications
                if sum(cls.object_counts.values()) > 8
            ]
            if len(high_density_slides) > 3:
                recommendations.append(f"Simplify slides with too much content: {high_density_slides}")
            
            # Recommendations based on consistency issues
            if consistency_issues:
                recommendations.append("Address consistency issues to improve presentation flow")
            
            # Recommendations based on slide types
            bullet_slide_count = sum(1 for cls in slide_classifications if cls.slide_type == SlideType.BULLET_SLIDE)
            if bullet_slide_count > len(slide_classifications) * 0.6:
                recommendations.append("Consider adding more visual elements to reduce text-heavy slides")
            
            return recommendations
            
        except Exception as e:
            logger.warning(f"Failed to generate recommendations: {e}")
            return []
    
    def _identify_strengths(
        self,
        slide_classifications: List[SlideClassification],
        structure: PresentationStructure
    ) -> List[str]:
        """Identify strengths of the presentation."""
        try:
            strengths = []
            
            # Good slide count
            if 5 <= structure.total_slides <= 20:
                strengths.append("Appropriate presentation length")
            
            # Good variety in slide types
            slide_types = set(cls.slide_type for cls in slide_classifications)
            if len(slide_types) >= 4:
                strengths.append("Good variety in slide types")
            
            # Presence of visual elements
            visual_slides = sum(1 for cls in slide_classifications 
                              if cls.slide_type in [SlideType.TABLE_SLIDE, SlideType.CHART_SLIDE, SlideType.IMAGE_SLIDE])
            if visual_slides > len(slide_classifications) * 0.3:
                strengths.append("Good use of visual elements")
            
            # Consistent structure
            if len(structure.structural_issues) == 0:
                strengths.append("Well-structured presentation")
            
            return strengths
            
        except Exception as e:
            logger.warning(f"Failed to identify strengths: {e}")
            return []
    
    def _identify_improvement_areas(
        self,
        slide_classifications: List[SlideClassification],
        structure: PresentationStructure,
        consistency_issues: List[str]
    ) -> List[str]:
        """Identify areas for improvement."""
        try:
            areas = []
            
            # Too many text-heavy slides
            text_heavy_slides = sum(1 for cls in slide_classifications if cls.slide_type == SlideType.BULLET_SLIDE)
            if text_heavy_slides > len(slide_classifications) * 0.7:
                areas.append("Reduce reliance on text-heavy slides")
            
            # Lack of visual elements
            visual_slides = sum(1 for cls in slide_classifications 
                              if cls.slide_type in [SlideType.TABLE_SLIDE, SlideType.CHART_SLIDE, SlideType.IMAGE_SLIDE])
            if visual_slides < len(slide_classifications) * 0.2:
                areas.append("Add more visual elements (charts, images, tables)")
            
            # Structural issues
            if structure.structural_issues:
                areas.append("Address structural issues")
            
            # Consistency issues
            if consistency_issues:
                areas.append("Improve consistency across slides")
            
            return areas
            
        except Exception as e:
            logger.warning(f"Failed to identify improvement areas: {e}")
            return []
    
    def _collect_sample_content(
        self,
        presentation_data: Dict[str, Any],
        analysis_depth: AnalysisDepth
    ) -> Dict[str, Any]:
        """Collect sample content from the presentation."""
        try:
            slides = presentation_data.get('slides', [])
            sample_content = {}
            
            # Sample titles
            titles = []
            for slide_data in slides[:5]:  # First 5 slides
                slide_info = slide_data.get('slide_info')
                if slide_info and slide_info.title:
                    titles.append(slide_info.title)
            sample_content['sample_titles'] = titles
            
            # Sample text content
            if analysis_depth in [AnalysisDepth.DETAILED, AnalysisDepth.COMPREHENSIVE]:
                text_samples = []
                for slide_data in slides[:3]:  # First 3 slides
                    slide_info = slide_data.get('slide_info')
                    if slide_info and slide_info.text_elements:
                        for elem in slide_info.text_elements[:2]:  # First 2 text elements per slide
                            content = elem.get('content_plain', '')
                            if content.strip():
                                preview = content[:100] + "..." if len(content) > 100 else content
                                text_samples.append(preview)
                sample_content['sample_text'] = text_samples
            
            return sample_content
            
        except Exception as e:
            logger.warning(f"Failed to collect sample content: {e}")
            return {}
    
    def clear_cache(self):
        """Clear the analysis cache."""
        self._analysis_cache.clear()
        logger.debug("Presentation analysis cache cleared")