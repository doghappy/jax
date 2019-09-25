using Newtonsoft.Json.Linq;
using System;
using System.IO;

namespace DogHappy.Jax.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("********************************");
            Console.WriteLine("Mytos Language Exporter");
            Console.WriteLine("********************************");
            Console.WriteLine();
            Console.WriteLine("Please input 1 or 2:");
            Console.WriteLine("1. Json files to Xlsx");
            Console.WriteLine("2. Xlsx to Json files");
            Console.WriteLine();

            if (int.TryParse(Console.ReadLine(), out int number))
            {
                switch (number)
                {
                    case 1:
                        JsonToXlsx();
                        break;
                    case 2:
                        XlsxToJson();
                        break;
                }
            }
            Console.WriteLine("Done");
        }

        static void JsonToXlsx()
        {
            var cvt = new JsonToXlsxConverter();
            string configJson = File.ReadAllText("config.json");
            var config = JObject.Parse(configJson);
            string masterJsonPath = config["JsonToXlsx"].Value<string>("MasterJson");
            string[] slaveJsonPaths = config["JsonToXlsx"]["SlaveJson"].ToObject<string[]>();
            string outputDir = config["JsonToXlsx"].Value<string>("OutputDir");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            using (var stream = cvt.GetXlsxFileAsync(masterJsonPath, slaveJsonPaths).GetAwaiter().GetResult())
            {
                string outputPath = Path.Combine(outputDir, "result.xlsx");
                File.WriteAllBytes(outputPath, stream.ToArray());
            }
        }

        static void XlsxToJson()
        {
            var cvt = new XlsxToJsonConverter();
            string configJson = File.ReadAllText("config.json");
            var config = JObject.Parse(configJson);
            string xlsxPath = config["XlsxToJson"].Value<string>("XlsxSource");
            string outputDir = config["XlsxToJson"].Value<string>("OutputDir");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            var list = cvt.GetJsonFilesAsync(xlsxPath).GetAwaiter().GetResult();
            foreach (var item in list)
            {
                string outputPath = Path.Combine(outputDir, DateTime.Now.ToString("yyyy-MM-ddThh_mm_ss") + item.Name);
                File.WriteAllBytes(outputPath, item.Stream.ToArray());
            }
        }
    }
}
