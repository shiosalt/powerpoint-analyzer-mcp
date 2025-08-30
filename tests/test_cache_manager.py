"""
Unit tests for CacheManager class.
"""

import pytest
import time
import tempfile
import os
from pathlib import Path
from unittest.mock import patch, MagicMock

from powerpoint_mcp_server.utils.cache_manager import (
    CacheManager, 
    get_global_cache, 
    reset_global_cache
)


class TestCacheManager:
    """Test cases for CacheManager class."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.cache = CacheManager(default_ttl=1, max_cache_size=3)
    
    def test_init(self):
        """Test CacheManager initialization."""
        cache = CacheManager(default_ttl=300, max_cache_size=50)
        assert cache.default_ttl == 300
        assert cache.max_cache_size == 50
        assert len(cache._cache) == 0
        assert len(cache._access_times) == 0
    
    def test_generate_file_hash_success(self):
        """Test successful file hash generation."""
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(b"test content")
            temp_path = temp_file.name
        
        try:
            hash1 = self.cache.generate_file_hash(temp_path)
            assert isinstance(hash1, str)
            assert len(hash1) == 64  # SHA256 hex length
            
            # Same file should generate same hash
            hash2 = self.cache.generate_file_hash(temp_path)
            assert hash1 == hash2
        finally:
            os.unlink(temp_path)
    
    def test_generate_file_hash_different_files(self):
        """Test that different files generate different hashes."""
        with tempfile.NamedTemporaryFile(delete=False) as temp_file1:
            temp_file1.write(b"content1")
            temp_path1 = temp_file1.name
        
        with tempfile.NamedTemporaryFile(delete=False) as temp_file2:
            temp_file2.write(b"content2")
            temp_path2 = temp_file2.name
        
        try:
            hash1 = self.cache.generate_file_hash(temp_path1)
            hash2 = self.cache.generate_file_hash(temp_path2)
            assert hash1 != hash2
        finally:
            os.unlink(temp_path1)
            os.unlink(temp_path2)
    
    def test_generate_file_hash_file_not_found(self):
        """Test file hash generation with non-existent file."""
        with pytest.raises(FileNotFoundError):
            self.cache.generate_file_hash("/non/existent/file.txt")
    
    def test_put_and_get_success(self):
        """Test successful cache put and get operations."""
        key = "test_key"
        data = {"test": "data"}
        
        # Put data in cache
        self.cache.put(key, data)
        
        # Retrieve data
        retrieved = self.cache.get(key)
        assert retrieved == data
    
    def test_get_nonexistent_key(self):
        """Test getting non-existent key returns None."""
        result = self.cache.get("nonexistent")
        assert result is None
    
    def test_cache_expiration(self):
        """Test that cache entries expire after TTL."""
        key = "expire_test"
        data = "test_data"
        
        # Put with very short TTL
        self.cache.put(key, data, ttl=0.1)
        
        # Should be available immediately
        assert self.cache.get(key) == data
        
        # Wait for expiration
        time.sleep(0.2)
        
        # Should be expired now
        assert self.cache.get(key) is None
    
    def test_cache_update_access_time(self):
        """Test that accessing cache updates access time."""
        key = "access_test"
        data = "test_data"
        
        self.cache.put(key, data)
        
        # Get initial access time
        initial_time = self.cache._access_times[key]
        
        # Wait a bit and access again
        time.sleep(0.01)
        self.cache.get(key)
        
        # Access time should be updated
        updated_time = self.cache._access_times[key]
        assert updated_time > initial_time
    
    def test_cache_size_limit_lru_eviction(self):
        """Test that cache evicts LRU entries when size limit is reached."""
        # Fill cache to capacity
        for i in range(3):
            self.cache.put(f"key_{i}", f"data_{i}")
        
        # Access key_1 to make it more recently used
        self.cache.get("key_1")
        
        # Add one more item (should evict key_0 as it's LRU)
        self.cache.put("key_3", "data_3")
        
        # key_0 should be evicted
        assert self.cache.get("key_0") is None
        # Others should still be there
        assert self.cache.get("key_1") == "data_1"
        assert self.cache.get("key_2") == "data_2"
        assert self.cache.get("key_3") == "data_3"
    
    def test_invalidate_existing_key(self):
        """Test invalidating existing cache entry."""
        key = "invalidate_test"
        data = "test_data"
        
        self.cache.put(key, data)
        assert self.cache.get(key) == data
        
        # Invalidate
        result = self.cache.invalidate(key)
        assert result is True
        assert self.cache.get(key) is None
    
    def test_invalidate_nonexistent_key(self):
        """Test invalidating non-existent key."""
        result = self.cache.invalidate("nonexistent")
        assert result is False
    
    def test_clear_cache(self):
        """Test clearing all cache entries."""
        # Add some entries
        for i in range(3):
            self.cache.put(f"key_{i}", f"data_{i}")
        
        assert len(self.cache._cache) == 3
        
        # Clear cache
        self.cache.clear()
        
        assert len(self.cache._cache) == 0
        assert len(self.cache._access_times) == 0
    
    def test_cleanup_expired(self):
        """Test cleanup of expired entries."""
        # Add entries with different TTLs
        self.cache.put("key_1", "data_1", ttl=0.1)  # Will expire
        self.cache.put("key_2", "data_2", ttl=10)   # Won't expire
        self.cache.put("key_3", "data_3", ttl=0.1)  # Will expire
        
        # Wait for some to expire
        time.sleep(0.2)
        
        # Cleanup expired
        removed_count = self.cache.cleanup_expired()
        
        assert removed_count == 2
        assert self.cache.get("key_1") is None
        assert self.cache.get("key_2") == "data_2"
        assert self.cache.get("key_3") is None
    
    def test_get_cache_stats(self):
        """Test cache statistics."""
        # Add some entries
        self.cache.put("key_1", "data_1", ttl=0.1)  # Will expire
        self.cache.put("key_2", "data_2", ttl=10)   # Won't expire
        
        # Wait for one to expire
        time.sleep(0.2)
        
        stats = self.cache.get_cache_stats()
        
        assert stats['total_entries'] == 2
        assert stats['expired_entries'] == 1
        assert stats['active_entries'] == 1
        assert stats['max_cache_size'] == 3
        assert stats['default_ttl'] == 1
    
    def test_thread_safety(self):
        """Test basic thread safety of cache operations."""
        import threading
        
        results = []
        errors = []
        
        def worker(worker_id):
            try:
                for i in range(10):
                    key = f"worker_{worker_id}_key_{i}"
                    data = f"worker_{worker_id}_data_{i}"
                    
                    self.cache.put(key, data)
                    retrieved = self.cache.get(key)
                    
                    if retrieved == data:
                        results.append(True)
                    else:
                        results.append(False)
            except Exception as e:
                errors.append(e)
        
        # Create multiple threads
        threads = []
        for i in range(3):
            thread = threading.Thread(target=worker, args=(i,))
            threads.append(thread)
            thread.start()
        
        # Wait for all threads to complete
        for thread in threads:
            thread.join()
        
        # Check results
        assert len(errors) == 0, f"Errors occurred: {errors}"
        assert all(results), "Some cache operations failed"


class TestGlobalCache:
    """Test cases for global cache functions."""
    
    def setup_method(self):
        """Reset global cache before each test."""
        reset_global_cache()
    
    def teardown_method(self):
        """Reset global cache after each test."""
        reset_global_cache()
    
    def test_get_global_cache_singleton(self):
        """Test that global cache returns same instance."""
        cache1 = get_global_cache()
        cache2 = get_global_cache()
        
        assert cache1 is cache2
        assert isinstance(cache1, CacheManager)
    
    def test_reset_global_cache(self):
        """Test resetting global cache."""
        cache1 = get_global_cache()
        cache1.put("test", "data")
        
        reset_global_cache()
        
        cache2 = get_global_cache()
        assert cache2 is not cache1
        assert cache2.get("test") is None


class TestCacheManagerIntegration:
    """Integration tests for CacheManager with file operations."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.cache = CacheManager()
    
    def test_file_hash_based_caching(self):
        """Test caching using file hash as key."""
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(b"test content for caching")
            temp_path = temp_file.name
        
        try:
            # Generate hash-based key
            cache_key = self.cache.generate_file_hash(temp_path)
            
            # Cache some processed data
            processed_data = {
                "file_path": temp_path,
                "processed_content": "extracted data",
                "metadata": {"size": 100, "type": "pptx"}
            }
            
            self.cache.put(cache_key, processed_data)
            
            # Retrieve using same file
            retrieved = self.cache.get(cache_key)
            assert retrieved == processed_data
            
            # Different file should have different key
            with tempfile.NamedTemporaryFile(delete=False) as temp_file2:
                temp_file2.write(b"different content")
                temp_path2 = temp_file2.name
            
            try:
                cache_key2 = self.cache.generate_file_hash(temp_path2)
                assert cache_key != cache_key2
                assert self.cache.get(cache_key2) is None
            finally:
                os.unlink(temp_path2)
                
        finally:
            os.unlink(temp_path)
    
    def test_file_modification_invalidates_cache(self):
        """Test that file modification changes hash and invalidates cache."""
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(b"initial content")
            temp_path = temp_file.name
        
        try:
            # Get initial hash and cache data
            initial_hash = self.cache.generate_file_hash(temp_path)
            self.cache.put(initial_hash, "initial_data")
            
            # Wait a bit to ensure different mtime
            time.sleep(0.01)
            
            # Modify file
            with open(temp_path, 'ab') as f:
                f.write(b" modified")
            
            # Hash should be different now
            modified_hash = self.cache.generate_file_hash(temp_path)
            assert modified_hash != initial_hash
            
            # Old cache entry should still exist but new hash won't find it
            assert self.cache.get(initial_hash) == "initial_data"
            assert self.cache.get(modified_hash) is None
            
        finally:
            os.unlink(temp_path)