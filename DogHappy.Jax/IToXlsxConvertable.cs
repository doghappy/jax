using System.IO;
using System.Threading.Tasks;

namespace DogHappy.Jax
{
    public interface IToXlsxConvertable
    {
        string SheetName { get; }
        Task<MemoryStream> GetXlsxFileAsync(string masterJson, params string[] slaveJsons);
        Task<MemoryStream> GetXlsxFileAsync(StreamName<Stream> masterJson, params StreamName<Stream>[] slaveJsons);
    }
}
