using System.Threading.Tasks;

namespace ExcelAnalyze.Logic
{
    /// <summary>
    /// Provides methods for interacting with Excel files.
    /// </summary>
    public interface IExcelFileService
    {
        /// <summary>
        /// Retrieves size details of the workbook located at the specified path.
        /// </summary>
        /// <param name="path">The file path to the Excel workbook.</param>
        /// <param name="password">An optional password required to open the workbook, if applicable.</param>
        /// <returns>
        /// A task representing the asynchronous operation, which returns
        /// a string containing size details of the workbook upon completion.
        /// </returns>
        Task<string> GetWorkBookSizesAsync(string path, string password);
    }
}