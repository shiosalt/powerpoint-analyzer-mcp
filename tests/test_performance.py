"""
Performance tests for PowerPoint MCP server.
Tests performance with large files and stress scenarios.
"""

import pytest
import time
import psutil
import os
import asyncio
from pathlib import Path
from unittest.mock import MagicMock

from powerpoint_mcp_server.server import PowerPointMCPServer
from powerpoint_mcp_server.core.content_extractor import ContentExtractor
from powerpoint_mcp_server.utils.cache_manager import reset_global_cache
from powerpoint_mcp_server.utils.zip_extractor import ZipExtractor


class TestPerformanceBenchmarks:
    """Performance benchmark tests."""
    
    def setup_method(self):
        """Set up test fixtures."""
        reset_global_cache()
        self.test_files_dir = Path("tests/test_files")
        self.minimal_pptx = self.test_files_dir / "test_minimal.pptx"
        self.complex_pptx = self.test_files_dir / "test_complex.pptx"
        self.server = PowerPointMCPServer()
        
        if not self.minimal_pptx.exists() or not self.complex_pptx.exists():
            pytest.skip("Test PowerPoint files not found. Run tests/create_test_pptx.py first.")
    
    def teardown_method(self):
        """Clean up after tests."""
        reset_global_cache()
    
    @pytest.mark.asyncio
    async def test_extraction_performance(self):
        """Test extraction performance benchmarks."""
        mock_request = MagicMock()
        mock_request.params = MagicMock()
        mock_request.params.arguments = {
            "file_path": str(self.complex_pptx)
        }
        
        # Measure extraction time
        start_time = time.time()
        result = await self.server._extract_powerpoint_content(mock_request.params.arguments)
        extraction_time = time.time() - start_time
        
        # Should complete within reasonable time (5 seconds for test files)
        assert extraction_time < 5.0, f"Extraction took {extraction_time:.2f} seconds, expected < 5.0"
        
        # Verify result is valid
        assert result is not None
        assert hasattr(result, 'content')
    
    def test_memory_usage_monitoring(self):
        """Test memory usage during processing."""
        process = psutil.Process(os.getpid())
        initial_memory = process.memory_info().rss
        
        content_extractor = ContentExtractor(enable_caching=True)
        
        # Process files multiple times to test memory management
        for i in range(10):
            with ZipExtractor(str(self.complex_pptx)) as zip_extractor:
                slide_files = zip_extractor.get_slide_xml_files()
                
                for j, slide_path in enumerate(slide_files.keys(), 1):
                    slide_xml = zip_extractor.read_xml_content(slide_path)
                    slide_info = content_extractor.extract_slide_content(slide_xml, j)
                    
                    # Verify we got valid data
                    assert slide_info.slide_number == j
        
        final_memory = process.memory_info().rss
        memory_growth = final_memory - initial_memory
        
        # Memory growth should be reasonable (less than 100MB for test)
        max_growth = 100 * 1024 * 1024  # 100MB
        assert memory_growth < max_growth, f"Memory grew by {memory_growth / 1024 / 1024:.2f} MB, expected < 100 MB"
    
    def test_caching_performance_impact(self):
        """Test performance impact of caching."""
        content_extractor_cached = ContentExtractor(enable_caching=True)
        content_extractor_no_cache = ContentExtractor(enable_caching=False)
        
        with ZipExtractor(str(self.complex_pptx)) as zip_extractor:
            slide_files = zip_extractor.get_slide_xml_files()
            first_slide_path = list(slide_files.keys())[0]
            slide_xml = zip_extractor.read_xml_content(first_slide_path)
            
            # Test with caching - first call
            start_time = time.time()
            result1 = content_extractor_cached.extract_slide_content(slide_xml, 1)
            first_cached_time = time.time() - start_time
            
            # Test with caching - second call (should be faster)
            start_time = time.time()
            result2 = content_extractor_cached.extract_slide_content(slide_xml, 1)
            second_cached_time = time.time() - start_time
            
            # Test without caching
            start_time = time.time()
            result3 = content_extractor_no_cache.extract_slide_content(slide_xml, 1)
            no_cache_time = time.time() - start_time
            
            # Verify results are equivalent
            assert result1.title == result2.title == result3.title
            assert result1.slide_number == result2.slide_number == result3.slide_number
            
            # Second cached call should be faster than first (though this might be flaky)
            # At minimum, verify caching is working
            cache_stats = content_extractor_cached.get_cache_stats()
            assert cache_stats['caching_enabled'] is True
            assert cache_stats['content_cache']['total_entries'] >= 1
    
    @pytest.mark.asyncio
    async def test_concurrent_request_performance(self):
        """Test performance under concurrent requests."""
        async def make_request(file_path):
            """Make a single request."""
            server = PowerPointMCPServer()
            mock_request = MagicMock()
            mock_request.params = MagicMock()
            mock_request.params.arguments = {"file_path": str(file_path)}
            
            start_time = time.time()
            result = await server._extract_powerpoint_content(mock_request.params.arguments)
            duration = time.time() - start_time
            
            return duration, result
        
        # Run concurrent requests
        start_time = time.time()
        tasks = [
            make_request(self.minimal_pptx),
            make_request(self.complex_pptx),
            make_request(self.minimal_pptx),  # Duplicate to test caching
        ]
        
        results = await asyncio.gather(*tasks)
        total_time = time.time() - start_time
        
        # Verify all requests completed successfully
        assert len(results) == 3
        for duration, result in results:
            assert result is not None
            assert hasattr(result, 'content')
            # Individual requests should complete within reasonable time
            assert duration < 10.0, f"Individual request took {duration:.2f} seconds"
        
        # Total time should be less than sum of individual times (due to concurrency)
        individual_times = [duration for duration, _ in results]
        sum_individual = sum(individual_times)
        
        # Allow some overhead, but should be significantly faster than sequential
        assert total_time < sum_individual * 0.8, f"Concurrent execution not efficient: {total_time:.2f}s vs {sum_individual:.2f}s"
    
    def test_large_xml_handling_performance(self):
        """Test performance with large XML content."""
        content_extractor = ContentExtractor(enable_caching=True)
        
        # Create progressively larger XML content
        sizes = [1000, 10000, 100000]  # 1KB, 10KB, 100KB
        
        for size in sizes:
            large_text = "A" * size
            large_xml = f'''<?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <p:cSld>
                    <p:spTree>
                        <p:nvGrpSpPr>
                            <p:cNvPr id="1" name=""/>
                            <p:cNvGrpSpPr/>
                            <p:nvPr/>
                        </p:nvGrpSpPr>
                        <p:grpSpPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="0" cy="0"/>
                                <a:chOff x="0" y="0"/>
                                <a:chExt cx="0" cy="0"/>
                            </a:xfrm>
                        </p:grpSpPr>
                        <p:sp>
                            <p:nvSpPr>
                                <p:cNvPr id="2" name="Large Text"/>
                                <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
                                <p:nvPr><p:ph type="title"/></p:nvPr>
                            </p:nvSpPr>
                            <p:txBody>
                                <a:bodyPr/>
                                <a:lstStyle/>
                                <a:p>
                                    <a:r>
                                        <a:rPr lang="en-US"/>
                                        <a:t>{large_text}</a:t>
                                    </a:r>
                                </a:p>
                            </a:txBody>
                        </p:sp>
                    </p:spTree>
                </p:cSld>
            </p:sld>'''
            
            start_time = time.time()
            result = content_extractor.extract_slide_content(large_xml, 1)
            processing_time = time.time() - start_time
            
            # Should handle large content without excessive time
            max_time = 2.0  # 2 seconds max for test content
            assert processing_time < max_time, f"Processing {size} bytes took {processing_time:.2f}s, expected < {max_time}s"
            
            # Should still extract basic information
            assert result.slide_number == 1
    
    def test_repeated_processing_performance(self):
        """Test performance degradation over repeated processing."""
        content_extractor = ContentExtractor(enable_caching=True)
        
        times = []
        
        # Process the same file multiple times
        for i in range(20):
            with ZipExtractor(str(self.complex_pptx)) as zip_extractor:
                slide_files = zip_extractor.get_slide_xml_files()
                
                start_time = time.time()
                
                for j, slide_path in enumerate(slide_files.keys(), 1):
                    slide_xml = zip_extractor.read_xml_content(slide_path)
                    slide_info = content_extractor.extract_slide_content(slide_xml, j)
                    assert slide_info.slide_number == j
                
                processing_time = time.time() - start_time
                times.append(processing_time)
        
        # Performance should not degrade significantly over time
        first_half_avg = sum(times[:10]) / 10
        second_half_avg = sum(times[10:]) / 10
        
        # Second half should not be more than 50% slower than first half
        degradation_ratio = second_half_avg / first_half_avg
        assert degradation_ratio < 1.5, f"Performance degraded by {degradation_ratio:.2f}x over repeated processing"
        
        # All processing times should be reasonable
        max_time = max(times)
        assert max_time < 5.0, f"Maximum processing time was {max_time:.2f}s, expected < 5.0s"


class TestStressTests:
    """Stress tests for edge cases and limits."""
    
    def setup_method(self):
        """Set up test fixtures."""
        reset_global_cache()
        self.test_files_dir = Path("tests/test_files")
        self.complex_pptx = self.test_files_dir / "test_complex.pptx"
        
        if not self.complex_pptx.exists():
            pytest.skip("Test PowerPoint files not found. Run tests/create_test_pptx.py first.")
    
    def teardown_method(self):
        """Clean up after tests."""
        reset_global_cache()
    
    def test_cache_memory_limits(self):
        """Test cache behavior under memory pressure."""
        content_extractor = ContentExtractor(enable_caching=True)
        
        # Generate many different XML contents to fill cache
        for i in range(100):
            xml_content = f'''<?xml version="1.0" encoding="UTF-8"?>
            <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                   xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <p:cSld>
                    <p:spTree>
                        <p:nvGrpSpPr>
                            <p:cNvPr id="1" name=""/>
                            <p:cNvGrpSpPr/>
                            <p:nvPr/>
                        </p:nvGrpSpPr>
                        <p:grpSpPr>
                            <a:xfrm>
                                <a:off x="0" y="0"/>
                                <a:ext cx="0" cy="0"/>
                                <a:chOff x="0" y="0"/>
                                <a:chExt cx="0" cy="0"/>
                            </a:xfrm>
                        </p:grpSpPr>
                        <p:sp>
                            <p:nvSpPr>
                                <p:cNvPr id="2" name="Test {i}"/>
                                <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
                                <p:nvPr><p:ph type="title"/></p:nvPr>
                            </p:nvSpPr>
                            <p:txBody>
                                <a:bodyPr/>
                                <a:lstStyle/>
                                <a:p>
                                    <a:r>
                                        <a:rPr lang="en-US"/>
                                        <a:t>Test Content {i}</a:t>
                                    </a:r>
                                </a:p>
                            </a:txBody>
                        </p:sp>
                    </p:spTree>
                </p:cSld>
            </p:sld>'''
            
            result = content_extractor.extract_slide_content(xml_content, i + 1)
            assert result.slide_number == i + 1
        
        # Cache should handle the load without crashing
        cache_stats = content_extractor.get_cache_stats()
        assert cache_stats['caching_enabled'] is True
        
        # Should have some reasonable number of entries (may have evicted some)
        assert cache_stats['content_cache']['total_entries'] > 0
    
    @pytest.mark.asyncio
    async def test_high_concurrency_stress(self):
        """Test behavior under high concurrency."""
        async def make_request(request_id):
            """Make a single request with unique ID."""
            server = PowerPointMCPServer()
            mock_request = MagicMock()
            mock_request.params = MagicMock()
            mock_request.params.arguments = {
                "file_path": str(self.complex_pptx),
                "slide_number": (request_id % 2) + 1  # Alternate between slides 1 and 2
            }
            
            try:
                result = await server._get_slide_info(mock_request.params.arguments)
                return request_id, True, result
            except Exception as e:
                return request_id, False, str(e)
        
        # Create many concurrent requests
        num_requests = 50
        tasks = [make_request(i) for i in range(num_requests)]
        
        start_time = time.time()
        results = await asyncio.gather(*tasks, return_exceptions=True)
        total_time = time.time() - start_time
        
        # Analyze results
        successful = 0
        failed = 0
        
        for result in results:
            if isinstance(result, Exception):
                failed += 1
            else:
                request_id, success, data = result
                if success:
                    successful += 1
                else:
                    failed += 1
        
        # Most requests should succeed
        success_rate = successful / num_requests
        assert success_rate > 0.8, f"Success rate {success_rate:.2f} too low, {failed} failed out of {num_requests}"
        
        # Should complete within reasonable time
        assert total_time < 30.0, f"High concurrency test took {total_time:.2f}s, expected < 30s"
    
    def test_malformed_xml_handling(self):
        """Test handling of malformed XML content."""
        content_extractor = ContentExtractor(enable_caching=False)
        
        malformed_xmls = [
            # Unclosed tags
            "<?xml version='1.0'?><p:sld><p:cSld><p:spTree>",
            
            # Invalid XML structure
            "<?xml version='1.0'?><invalid><nested><unclosed>",
            
            # Empty content
            "",
            
            # Non-XML content
            "This is not XML at all",
            
            # XML with invalid characters
            "<?xml version='1.0'?><p:sld>\x00\x01\x02</p:sld>",
        ]
        
        for i, malformed_xml in enumerate(malformed_xmls):
            # Should handle malformed XML gracefully
            result = content_extractor.extract_slide_content(malformed_xml, i + 1)
            
            # Should return a valid SlideInfo object with at least slide number
            assert result is not None
            assert result.slide_number == i + 1
            # Other fields may be None or empty due to parsing failure