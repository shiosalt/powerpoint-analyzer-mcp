"""Slide selection utility for parsing Python-style slice notation."""

import re
from typing import List, Union, Optional, Any
import logging

logger = logging.getLogger(__name__)


def parse_slide_numbers(slide_spec: Union[str, List[int], int, None], total_slides: int) -> List[int]:
    """
    Parse slide numbers specification with Python-style slicing support.
    
    Args:
        slide_spec: Slide specification in various formats:
            - None: All slides
            - int: Single slide (e.g., 3)
            - List[int]: Specific slides (e.g., [1, 5, 10])
            - str: Python-style slicing:
                - ":100" or "[:100]": First 100 slides (1-100)
                - "5:20" or "[5:20]": Slides 5-20
                - "25:" or "[25:]": Slides 25 to end
                - "3" or "[3]": Single slide 3
                - "1,5,10" or "[1,5,10]": Specific slides 1, 5, 10
        total_slides: Total number of slides in presentation
        
    Returns:
        List[int]: List of slide numbers (1-based indexing)
        
    Raises:
        ValueError: If slide specification is invalid
        
    Examples:
        parse_slide_numbers(None, 100) -> [1, 2, ..., 100]
        parse_slide_numbers(3, 100) -> [3]
        parse_slide_numbers([1, 5, 10], 100) -> [1, 5, 10]
        parse_slide_numbers(":10", 100) -> [1, 2, ..., 10]
        parse_slide_numbers("5:20", 100) -> [5, 6, ..., 20]
        parse_slide_numbers("25:", 100) -> [25, 26, ..., 100]
        parse_slide_numbers("1,5,10", 100) -> [1, 5, 10]
    """
    if slide_spec is None:
        # Return all slides
        return list(range(1, total_slides + 1))
    
    if isinstance(slide_spec, int):
        # Single slide number
        if slide_spec < 1 or slide_spec > total_slides:
            raise ValueError(f"Slide number {slide_spec} is out of range (1-{total_slides})")
        return [slide_spec]
    
    if isinstance(slide_spec, list):
        # List of specific slide numbers (existing format)
        if not all(isinstance(x, int) for x in slide_spec):
            raise ValueError("All slide numbers must be integers")
        
        # Validate slide numbers
        for slide_num in slide_spec:
            if slide_num < 1 or slide_num > total_slides:
                raise ValueError(f"Slide number {slide_num} is out of range (1-{total_slides})")
        
        return sorted(list(set(slide_spec)))  # Remove duplicates and sort
    
    if isinstance(slide_spec, str):
        return _parse_string_slide_spec(slide_spec, total_slides)
    
    raise ValueError(f"Invalid slide specification type: {type(slide_spec)}")


def _parse_string_slide_spec(slide_spec: str, total_slides: int) -> List[int]:
    """Parse string-based slide specifications."""
    # Remove whitespace and optional brackets
    spec = slide_spec.strip()
    if spec.startswith('[') and spec.endswith(']'):
        spec = spec[1:-1].strip()
    
    # Check if it's a comma-separated list
    if ',' in spec and ':' not in spec:
        return _parse_comma_separated(spec, total_slides)
    
    # Check if it's a slice notation
    if ':' in spec:
        return _parse_slice_notation(spec, total_slides)
    
    # Single number as string
    try:
        slide_num = int(spec)
        if slide_num < 1 or slide_num > total_slides:
            raise ValueError(f"Slide number {slide_num} is out of range (1-{total_slides})")
        return [slide_num]
    except ValueError as e:
        if "out of range" in str(e):
            raise
        raise ValueError(f"Invalid slide specification: '{slide_spec}'")


def _parse_comma_separated(spec: str, total_slides: int) -> List[int]:
    """Parse comma-separated slide numbers."""
    try:
        slide_numbers = []
        for part in spec.split(','):
            part = part.strip()
            if not part:
                continue
            slide_num = int(part)
            if slide_num < 1 or slide_num > total_slides:
                raise ValueError(f"Slide number {slide_num} is out of range (1-{total_slides})")
            slide_numbers.append(slide_num)
        
        return sorted(list(set(slide_numbers)))  # Remove duplicates and sort
    except ValueError as e:
        if "out of range" in str(e):
            raise
        raise ValueError(f"Invalid comma-separated slide specification: '{spec}'")


def _parse_slice_notation(spec: str, total_slides: int) -> List[int]:
    """Parse Python-style slice notation."""
    # Split on ':'
    parts = spec.split(':')
    if len(parts) != 2:
        raise ValueError(f"Invalid slice notation: '{spec}'. Expected format: 'start:end'")
    
    start_str, end_str = parts
    start_str = start_str.strip()
    end_str = end_str.strip()
    
    # Parse start
    if start_str == '':
        start = 1  # Default to first slide
    else:
        try:
            start = int(start_str)
            if start < 1:
                raise ValueError(f"Start slide number must be >= 1, got {start}")
        except ValueError as e:
            if "must be >=" in str(e):
                raise
            raise ValueError(f"Invalid start slide number: '{start_str}'")
    
    # Parse end
    if end_str == '':
        end = total_slides  # Default to last slide
    else:
        try:
            end = int(end_str)
            if end < 1:
                raise ValueError(f"End slide number must be >= 1, got {end}")
        except ValueError as e:
            if "must be >=" in str(e):
                raise
            raise ValueError(f"Invalid end slide number: '{end_str}'")
    
    # Validate range
    if start > total_slides:
        raise ValueError(f"Start slide {start} is beyond total slides ({total_slides})")
    
    if end > total_slides:
        logger.warning(f"End slide {end} is beyond total slides ({total_slides}), capping to {total_slides}")
        end = total_slides
    
    if start > end:
        raise ValueError(f"Start slide ({start}) cannot be greater than end slide ({end})")
    
    return list(range(start, end + 1))


def validate_slide_numbers(slide_numbers: List[int], total_slides: int) -> List[int]:
    """
    Validate and filter slide numbers to ensure they're within valid range.
    
    Args:
        slide_numbers: List of slide numbers to validate
        total_slides: Total number of slides in presentation
        
    Returns:
        List[int]: Validated slide numbers within range
        
    Raises:
        ValueError: If no valid slide numbers remain after filtering
    """
    if not slide_numbers:
        return []
    
    valid_slides = [num for num in slide_numbers if 1 <= num <= total_slides]
    invalid_slides = [num for num in slide_numbers if num < 1 or num > total_slides]
    
    if invalid_slides:
        logger.warning(f"Ignoring invalid slide numbers: {invalid_slides} (valid range: 1-{total_slides})")
    
    if not valid_slides:
        raise ValueError(f"No valid slide numbers found. Valid range: 1-{total_slides}")
    
    return sorted(list(set(valid_slides)))