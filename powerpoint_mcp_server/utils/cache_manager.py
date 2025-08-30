"""
Cache manager for PowerPoint MCP server responses.
Provides file hash-based caching with memory storage and expiration.
"""

import hashlib
import time
from typing import Any, Dict, Optional, Tuple
from pathlib import Path
import threading
import weakref


class CacheManager:
    """
    Manages caching of PowerPoint processing results using file hash-based keys.
    Provides memory-based temporary caching with expiration.
    """
    
    def __init__(self, default_ttl: int = 3600, max_cache_size: int = 100):
        """
        Initialize the cache manager.
        
        Args:
            default_ttl: Default time-to-live for cache entries in seconds (default: 1 hour)
            max_cache_size: Maximum number of entries to keep in cache
        """
        self.default_ttl = default_ttl
        self.max_cache_size = max_cache_size
        self._cache: Dict[str, Tuple[Any, float]] = {}  # key -> (data, expiry_time)
        self._access_times: Dict[str, float] = {}  # key -> last_access_time
        self._lock = threading.RLock()
        
        # Weak reference to self for cleanup
        self._cleanup_ref = weakref.ref(self, self._cleanup_cache)
    
    def generate_file_hash(self, file_path: str) -> str:
        """
        Generate a hash-based cache key for a file.
        Uses file path, size, and modification time for uniqueness.
        
        Args:
            file_path: Path to the file
            
        Returns:
            Hash string to use as cache key
            
        Raises:
            FileNotFoundError: If file doesn't exist
            OSError: If file cannot be accessed
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        try:
            stat = path.stat()
            # Create hash from file path, size, and modification time
            hash_input = f"{file_path}:{stat.st_size}:{stat.st_mtime}"
            return hashlib.sha256(hash_input.encode('utf-8')).hexdigest()
        except OSError as e:
            raise OSError(f"Cannot access file {file_path}: {e}")
    
    def get(self, key: str) -> Optional[Any]:
        """
        Retrieve data from cache.
        
        Args:
            key: Cache key
            
        Returns:
            Cached data if found and not expired, None otherwise
        """
        with self._lock:
            if key not in self._cache:
                return None
            
            data, expiry_time = self._cache[key]
            current_time = time.time()
            
            # Check if expired
            if current_time > expiry_time:
                self._remove_entry(key)
                return None
            
            # Update access time
            self._access_times[key] = current_time
            return data
    
    def put(self, key: str, data: Any, ttl: Optional[int] = None) -> None:
        """
        Store data in cache.
        
        Args:
            key: Cache key
            data: Data to cache
            ttl: Time-to-live in seconds (uses default if None)
        """
        if ttl is None:
            ttl = self.default_ttl
        
        expiry_time = time.time() + ttl
        
        with self._lock:
            # Check if we need to make room
            if len(self._cache) >= self.max_cache_size and key not in self._cache:
                self._evict_lru()
            
            self._cache[key] = (data, expiry_time)
            self._access_times[key] = time.time()
    
    def invalidate(self, key: str) -> bool:
        """
        Remove specific entry from cache.
        
        Args:
            key: Cache key to remove
            
        Returns:
            True if entry was removed, False if not found
        """
        with self._lock:
            if key in self._cache:
                self._remove_entry(key)
                return True
            return False
    
    def clear(self) -> None:
        """Clear all cache entries."""
        with self._lock:
            self._cache.clear()
            self._access_times.clear()
    
    def cleanup_expired(self) -> int:
        """
        Remove all expired entries from cache.
        
        Returns:
            Number of entries removed
        """
        current_time = time.time()
        expired_keys = []
        
        with self._lock:
            for key, (_, expiry_time) in self._cache.items():
                if current_time > expiry_time:
                    expired_keys.append(key)
            
            for key in expired_keys:
                self._remove_entry(key)
        
        return len(expired_keys)
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """
        Get cache statistics.
        
        Returns:
            Dictionary with cache statistics
        """
        with self._lock:
            current_time = time.time()
            expired_count = sum(1 for _, expiry_time in self._cache.values() 
                              if current_time > expiry_time)
            
            return {
                'total_entries': len(self._cache),
                'expired_entries': expired_count,
                'active_entries': len(self._cache) - expired_count,
                'max_cache_size': self.max_cache_size,
                'default_ttl': self.default_ttl
            }
    
    def _remove_entry(self, key: str) -> None:
        """Remove entry from both cache and access times."""
        self._cache.pop(key, None)
        self._access_times.pop(key, None)
    
    def _evict_lru(self) -> None:
        """Evict least recently used entry to make room."""
        if not self._access_times:
            return
        
        # Find least recently used key
        lru_key = min(self._access_times.keys(), 
                     key=lambda k: self._access_times[k])
        self._remove_entry(lru_key)
    
    @staticmethod
    def _cleanup_cache(cache_ref):
        """Cleanup method called when CacheManager is garbage collected."""
        # This is called when the CacheManager is being destroyed
        # No cleanup needed for memory-based cache
        pass


# Global cache instance
_global_cache = None
_cache_lock = threading.Lock()


def get_global_cache() -> CacheManager:
    """
    Get the global cache manager instance.
    Creates one if it doesn't exist.
    
    Returns:
        Global CacheManager instance
    """
    global _global_cache
    
    if _global_cache is None:
        with _cache_lock:
            if _global_cache is None:
                _global_cache = CacheManager()
    
    return _global_cache


def reset_global_cache() -> None:
    """Reset the global cache instance (mainly for testing)."""
    global _global_cache
    
    with _cache_lock:
        if _global_cache is not None:
            _global_cache.clear()
        _global_cache = None