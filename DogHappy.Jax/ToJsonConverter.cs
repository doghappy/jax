using Newtonsoft.Json.Linq;
using System.Data;
using System.IO;
using System.Linq;

namespace DogHappy.Jax
{
    public abstract class ToJsonConverter
    {
        public abstract char Separator { get; set; }

        protected void WriteJson(JObject jobj, DataRow row, string col, string key)
        {
            int index = key.IndexOf(Separator);
            if (index == -1)
            {
                jobj.Add(key, new JValue(row[col].ToString()));
            }
            else
            {
                var keys = key.Split(Separator);
                foreach (var item in keys)
                {
                    var last = keys.Last();
                    if (last != item)
                    {
                        string newKey = key.Substring(index + 1);
                        if (jobj.ContainsKey(item))
                        {
                            WriteJson(jobj.Value<JObject>(item), row, col, newKey);
                            return;
                        }
                        else
                        {
                            var subObject = new JObject();
                            jobj.Add(item, subObject);
                            WriteJson(subObject, row, col, newKey);
                            return;
                        }
                    }
                }
            }
        }
    }
}
