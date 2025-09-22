using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Skill_Loop
{
    // OcrCacheService.cs
    // OcrCacheService.cs
    public static class OcrCacheService
    {
        private static readonly ConcurrentDictionary<string, CacheItem> _cache = new ConcurrentDictionary<string, CacheItem>();
        private static readonly TimeSpan _defaultCacheDuration = TimeSpan.FromSeconds(5);

        public static string GetOrAdd(string key, Func<string> valueFactory, TimeSpan? duration = null)
        {
            var cacheDuration = duration ?? _defaultCacheDuration;

            if (_cache.TryGetValue(key, out var item) &&
                DateTime.Now - item.Timestamp < cacheDuration)
            {
                return item.Value;
            }

            string value = valueFactory();
            _cache[key] = new CacheItem { Value = value, Timestamp = DateTime.Now };
            return value;
        }

        public static void Clear()
        {
            _cache.Clear();
        }

        public static void RemoveExpiredItems()
        {
            var expiredKeys = _cache
                .Where(kv => DateTime.Now - kv.Value.Timestamp > _defaultCacheDuration)
                .Select(kv => kv.Key)
                .ToList();

            foreach (var key in expiredKeys)
            {
                _cache.TryRemove(key, out _);
            }
        }

        private class CacheItem
        {
            public string Value { get; set; }
            public DateTime Timestamp { get; set; }
        }
    }
}
