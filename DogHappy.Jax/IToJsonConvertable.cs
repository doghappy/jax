using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace DogHappy.Jax
{
    public interface IToJsonConvertable
    {
        int SheetIndex { get; }
        Task<List<StreamName<MemoryStream>>> GetJsonFilesAsync(string xlsxPath);
        Task<List<StreamName<MemoryStream>>> GetJsonFilesAsync(Stream xlsx);
    }
}
