using ExcelDataReader;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace DogHappy.Jax
{
    public class XlsxToJsonConverter : ToJsonConverter, IToJsonConvertable
    {
        public int SheetIndex { get; set; }

        public override char Separator { get; set; } = '.';

        public Task<List<StreamName<MemoryStream>>> GetJsonFilesAsync(string xlsxPath)
        {
            using (var stream = File.Open(xlsxPath, FileMode.Open, FileAccess.Read))
                return GetJsonFilesAsync(stream);
        }

        public Task<List<StreamName<MemoryStream>>> GetJsonFilesAsync(Stream xlsx)
        {
            var list = new List<StreamName<MemoryStream>>();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (var reader = ExcelReaderFactory.CreateReader(xlsx))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = config => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                });
                var table = result.Tables[SheetIndex];
                foreach (DataColumn col in table.Columns)
                {
                    if (table.Columns.IndexOf(col) > 0)
                    {
                        var jobj = new JObject();
                        foreach (DataRow row in table.Rows)
                        {
                            string key = row["key"].ToString();
                            WriteJson(jobj, row, col.ColumnName, key);
                        }
                        list.Add(new StreamName<MemoryStream>
                        {
                            Stream = new MemoryStream(Encoding.UTF8.GetBytes(jobj.ToString())),
                            Name = col.ColumnName
                        });
                    }
                }
            }
            return Task.FromResult(list);
        }
    }
}
