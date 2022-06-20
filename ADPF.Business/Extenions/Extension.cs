using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADPF.Business.Extenions
{
    public static class Extension
    {
        public static string DumpJson<T>(this T o)
        {
            string Json = JsonConvert.SerializeObject(o, Formatting.Indented);
            return Json;
        }
    }
}
