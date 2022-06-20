using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Business.BHelpers
{
    public class MemoryCacheHelper
    {
        public static object GetValue(string key)
        {
            MemoryCache memoryCache = MemoryCache.Default;
            return memoryCache.Get(key);
        }

        public static bool Add(string key, object value, int durationHour)
        {
            MemoryCache memoryCache = MemoryCache.Default;
            return memoryCache.Add(key, value, DateTimeOffset.UtcNow.AddHours(durationHour));
        }

        public static void Delete(string key)
        {
            MemoryCache memoryCache = MemoryCache.Default;
            if (memoryCache.Contains(key))
            {
                memoryCache.Remove(key);
            }
        }
    }
}
