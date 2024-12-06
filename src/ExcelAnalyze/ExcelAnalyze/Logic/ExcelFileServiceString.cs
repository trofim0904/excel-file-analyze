using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAnalyze.Model;
using Excel = Microsoft.Office.Interop.Excel;
using Message = ExcelAnalyze.Descriptor.Message;

namespace ExcelAnalyze.Logic
{
    public class ExcelFileServiceString : IExcelFileService
    {
        public async Task<string> GetWorkBookSizesAsync(string path, string password)
        {
            var result = CreateBaseStringBuilder();
            try
            {
                long fileBytes = new System.IO.FileInfo(path).Length;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = null;
                try
                {
                    workbook = GetWorkbook(path, password, excelApp);
                    var worksheetTasks = new List<Task<Tuple<string, long>>>();
                    foreach (Excel.Worksheet worksheet in workbook.Sheets)
                    {
                        worksheetTasks.Add(ProcessWorksheet(worksheet));
                    }
                    var worksheetsInfo = await Task.WhenAll(worksheetTasks);
                    var totalUsedRange = worksheetsInfo.Sum(info => info.Item2);
                    foreach (var worksheetInfo in worksheetsInfo.OrderByDescending(i => i.Item2))
                    {
                        var worksheet = new Worksheet(worksheetInfo.Item1, 
                            worksheetInfo.Item2, totalUsedRange, fileBytes);
                        result.AppendLine(worksheet.ToString());
                    }
                    AddFooterInfo(result);
                    return result.ToString();
                }
                catch (Exception ex)
                {
                    Helper.ShowError(ex);
                }
                finally
                {
                    workbook?.Close(false);
                    excelApp.Quit();
                }
            }
            catch (Exception ex)
            {
                Helper.ShowError(ex);
            }
            return string.Concat(Message.ProcessFinished, DateTime.Now);
        }

        private static void AddFooterInfo(StringBuilder result)
        {
            result.AppendLine(Message.Line);
            result.AppendLine();
            result.AppendLine(string.Concat(Message.ProcessFinished, DateTime.Now));
            result.AppendLine(Message.Note);
        }

        private static Excel.Workbook GetWorkbook(string path, string password, Excel.Application excelApp)
        {
            var workbook = string.IsNullOrEmpty(password)
                ? excelApp.Workbooks.Open(path, ReadOnly: true)
                : excelApp.Workbooks.Open(path, Password: password, ReadOnly: true);
            return workbook;
        }

        private static StringBuilder CreateBaseStringBuilder()
        {
            StringBuilder result = new StringBuilder();
            result.AppendLine(string.Concat(Message.ProcessStartedTime, DateTime.Now));
            result.AppendLine();
            result.AppendLine(Message.Worksheets);
            result.AppendLine(Message.Line);
            return result;
        }

        private Task<Tuple<string, long>> ProcessWorksheet(Excel.Worksheet worksheet)
        {
            Excel.Range usedRange = worksheet.UsedRange;
            return Task.FromResult(new Tuple<string, long>(worksheet.Name, usedRange.Rows.Count));
        }
    }
}