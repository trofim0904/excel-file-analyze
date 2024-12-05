namespace ExcelAnalyze.Logic
{
    public interface IExcelFileService
    {
        string GetWorkBookSizes(string path, string password);
    }
}