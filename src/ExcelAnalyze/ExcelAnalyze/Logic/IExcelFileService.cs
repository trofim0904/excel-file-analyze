using System.Threading.Tasks;

namespace ExcelAnalyze.Logic
{
    public interface IExcelFileService
    {
        Task<string> GetWorkBookSizesAsync(string path, string password);
    }
}