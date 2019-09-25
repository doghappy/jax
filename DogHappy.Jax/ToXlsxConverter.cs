using System;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Linq;

namespace DogHappy.Jax
{
    public abstract class ToXlsxConverter
    {
        public abstract char Separator { get; set; }

        protected void ReadMasterJson(JObject jObj, DataTable table, string prefix = "", int objLv = 0)
        {
            objLv++;
            var props = jObj.Properties();
            foreach (var item in props)
            {
                if (item.HasValues)
                {
                    if (item.Type == JTokenType.Property)
                    {
                        var property = item as JProperty;
                        if (property.Name.Contains(Separator))
                        {
                            throw new ArgumentException("json property name it not allowed to contains Separator, default Separator is '.', maybe need to change the 'Converter.Separator'");
                        }

                        if (property.Value.Type != JTokenType.Object && property.Value.Type != JTokenType.Property)
                        {
                            string key = prefix + property.Name;
                            table.Rows.Add(key, property.Value.ToString());
                        }
                        else if (property.Value.Type == JTokenType.Object)
                        {
                            string newPrefix = null;
                            if (objLv == 1)
                            {
                                newPrefix = property.Name + Separator;
                            }
                            else
                            {
                                newPrefix = prefix + property.Name + Separator;
                            }
                            ReadMasterJson(property.Value as JObject, table, newPrefix, objLv);
                        }
                    }
                }
            }
        }

        protected void ReadSlaveJson(string col, JObject jObj, DataTable table, string prefix = "", int objLv = 0)
        {
            objLv++;
            var props = jObj.Properties();
            foreach (var item in props)
            {
                if (item.HasValues)
                {
                    if (item.Type == JTokenType.Property)
                    {
                        var property = item as JProperty;
                        if (property.Name.Contains(Separator))
                        {
                            throw new ArgumentException("json property name it not allowed to contains Separator, default Separator is '.', maybe need to change the 'Converter.Separator'");
                        }

                        if (property.Value.Type != JTokenType.Object && property.Value.Type != JTokenType.Property)
                        {
                            string key = prefix + property.Name;
                            var row = table.Rows.Find(key);
                            if (row != null)
                            {
                                row[col] = property.Value.ToString();
                            }
                        }
                        else if (property.Value.Type == JTokenType.Object)
                        {
                            string newPrefix = null;
                            if (objLv == 1)
                            {
                                newPrefix = property.Name + Separator;
                            }
                            else
                            {
                                newPrefix = prefix + property.Name + Separator;
                            }
                            ReadSlaveJson(col, property.Value as JObject, table, newPrefix, objLv);
                        }
                    }
                }
            }
        }
    }
}
