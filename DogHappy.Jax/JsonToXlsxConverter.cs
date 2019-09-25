using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System.Data;
using System.IO;
using System.Threading.Tasks;

namespace DogHappy.Jax
{
    public class JsonToXlsxConverter : ToXlsxConverter, IToXlsxConvertable
    {
        public string SheetName { get; set; } = "No Sheet Name";
        public override char Separator { get; set; } = '.';

        public async Task<MemoryStream> GetXlsxFileAsync(string masterJson, params string[] slaveJsons)
        {
            var masterStream = File.Open(masterJson, FileMode.Open, FileAccess.Read);
            var streams = new StreamName<Stream>[slaveJsons.Length];
            for (int i = 0; i < slaveJsons.Length; i++)
            {
                streams[i] = new StreamName<Stream>
                {
                    Stream = File.Open(slaveJsons[i], FileMode.Open, FileAccess.Read),
                    Name = Path.GetFileName(slaveJsons[i])
                };
            }
            try
            {
                return await GetXlsxFileAsync(new StreamName<Stream>
                {
                    Stream = masterStream,
                    Name = Path.GetFileName(masterJson)
                }, streams);
            }
            finally
            {
                masterStream.Dispose();
                for (int i = 0; i < streams.Length; i++)
                {
                    streams[i].Stream.Dispose();
                }
            }
        }


        public async Task<MemoryStream> GetXlsxFileAsync(StreamName<Stream> masterJson, params StreamName<Stream>[] slaveJsons)
        {
            using (DataTable table = new DataTable(SheetName))
            {
                DataColumn keyCol = table.Columns.Add("key");
                table.PrimaryKey = new[] { keyCol };
                table.Columns.Add(masterJson.Name);

                using (var reader = new StreamReader(masterJson.Stream))
                {
                    string json = await reader.ReadToEndAsync();
                    var jobj = JObject.Parse(json);
                    ReadMasterJson(jobj, table);
                }

                for (int i = 0; i < slaveJsons.Length; i++)
                {
                    string colName = slaveJsons[i].Name;
                    table.Columns.Add(colName);
                    using (var reader = new StreamReader(slaveJsons[i].Stream))
                    {
                        string json = await reader.ReadToEndAsync();
                        var jobj = JObject.Parse(json);
                        ReadSlaveJson(colName, jobj, table);
                    }
                }

                var memoryStream = new MemoryStream();
                using (var workbook = new XLWorkbook())
                {
                    workbook.Worksheets.Add(table);
                    workbook.SaveAs(memoryStream);
                }
                return memoryStream;
            }
        }
    }
}
