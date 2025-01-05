using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelAnalyze.Model;
using Excel = Microsoft.Office.Interop.Excel;
using Message = ExcelAnalyze.Descriptor.Message;

namespace ExcelAnalyze.Logic
{
    public class ExcelFileServiceInterop : ExcelFileServiceBase, IExcelFileService
    {
        public async Task<string> GetWorkBookSizesAsync(string path, string password)
        {
            var result = CreateBaseStringBuilder();
            try
            {
                var fileBytes = new FileInfo(path).Length;
                AddFileSizeInfo(result, fileBytes, path);
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = null;
                try
                {
                    result.AppendLine(Message.Line);
                    workbook = GetWorkbook(path, password, excelApp);
                    var worksheetTasks = new List<Task<Tuple<string, decimal>>>();
                    foreach (Excel.Worksheet worksheet in workbook.Sheets)
                    {
                        worksheetTasks.Add(ProcessWorksheet(worksheet));
                    }
                    var worksheetsInfo = await Task.WhenAll(worksheetTasks);
                    var totalWeight = worksheetsInfo.Sum(info => info.Item2);
                    foreach (var worksheetInfo in worksheetsInfo.OrderByDescending(i => i.Item2))
                    {
                        var worksheet = new Worksheet(worksheetInfo.Item1, worksheetInfo.Item2, totalWeight, fileBytes);
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

        private static Excel.Workbook GetWorkbook(string path, string password, Excel.Application excelApp)
        {
            var workbook = string.IsNullOrEmpty(password)
                ? excelApp.Workbooks.Open(path, ReadOnly: true)
                : excelApp.Workbooks.Open(path, Password: password, ReadOnly: true);
            return workbook;
        }

        private Task<Tuple<string, decimal>> ProcessWorksheet(Excel.Worksheet worksheet)
        {
            return Task.FromResult(new Tuple<string, decimal>(worksheet.Name, CalculateWorksheetWeight(worksheet)));
        }

        private decimal CalculateWorksheetWeight(Excel.Worksheet worksheet)
        {
            // Estimated weight per element
            // Average weight per comment
            const decimal weightPerComment = 8;
            // Average weight per hyperlink
            const decimal weightPerHyperlink = 3;
            // Average weight per shape
            const decimal weightPerShape = 39;
            // Average weight per cell
            const decimal weightPerCell = 1;
            return worksheet.Comments.Count * weightPerComment +
                   worksheet.Hyperlinks.Count * weightPerHyperlink +
                   worksheet.Shapes.Count * weightPerShape +
                   worksheet.UsedRange.Count * weightPerCell;
        }
    }
}