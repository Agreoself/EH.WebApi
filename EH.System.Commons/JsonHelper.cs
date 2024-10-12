using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Commons
{
    public static class JsonHelper
    {
        public static string ToJson(this object obj)
        {
            if (obj == null)
                return "";
            return JsonConvert.SerializeObject(obj, Formatting.Indented);
        }

        public static T ToObject<T>(this object obj)
        {
            if (obj == null)
                return default(T);
            return JsonConvert.DeserializeObject<T>(obj.ToJson());
        }
    }
}
