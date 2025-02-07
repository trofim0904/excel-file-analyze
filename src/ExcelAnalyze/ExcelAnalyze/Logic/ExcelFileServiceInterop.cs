using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Message = ExcelAnalyze.Descriptor.Message;

namespace ExcelAnalyze.Logic
{
    public class ExcelFileServiceInterop : ExcelFileServiceBase, IExcelFileService
    {
        public Task<string> GetWorkBookSizesAsync(string path, string password)
        {
            return Task.Run(() => GetWorkBookSizesAsyncProcess(path, password));
        }

        /// <summary>
        /// Executes an asynchronous process to retrieve workbook size details from the specified file.
        /// </summary>
        /// <param name="path">The file path to the workbook.</param>
        /// <param name="password">The password required to open the workbook.</param>
        /// <returns>
        /// A <see cref="Task{TResult}"/> that, when completed, returns a string containing workbook size details.
        /// </returns>
        private async Task<string> GetWorkBookSizesAsyncProcess(string path, string password)
        {
            var result = CreateBaseStringBuilder();
            try
            {
                var fileBytes = new FileInfo(path).Length;
                AddFileSizeInfo(result, fileBytes, path);
                Application excelApp = new Application();
                Workbook workbook = null;
                try
                {
                    result.AppendLine(Message.Line);
                    workbook = GetWorkbook(path, password, excelApp);
                    var worksheetTasks = new List<Task<Tuple<string, decimal>>>();
                    foreach (Worksheet worksheet in workbook.Sheets)
                    {
                        worksheetTasks.Add(ProcessWorksheet(worksheet));
                    }
                    var worksheetsInfo = await Task.WhenAll(worksheetTasks);
                    var totalWeight = worksheetsInfo.Sum(info => info.Item2);
                    foreach (var worksheetInfo in worksheetsInfo.OrderByDescending(i => i.Item2))
                    {
                        var worksheet = new Model.Worksheet(worksheetInfo, totalWeight, fileBytes);
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
                    if (workbook != null)
                    {
                        workbook.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    }
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                Helper.ShowError(ex);
            }
            return string.Concat(Message.ProcessFinished, DateTime.Now);
        }

        /// <summary>
        /// Opens an Excel workbook at the specified path, optionally using the provided password.
        /// </summary>
        /// <param name="path">The file path to the Excel workbook.</param>
        /// <param name="password">An optional password for protected workbooks.</param>
        /// <param name="excelApp">The current <see cref="Excel.Application"/> instance.</param>
        /// <returns>
        /// An <see cref="Workbook"/> object representing the opened workbook.
        /// </returns>
        private static Workbook GetWorkbook(string path, string password, Application excelApp)
        {
            var workbook = string.IsNullOrEmpty(password)
                ? excelApp.Workbooks.Open(path, ReadOnly: true)
                : excelApp.Workbooks.Open(path, Password: password, ReadOnly: true);
            return workbook;
        }


        /// <summary>
        /// Asynchronously processes the specified Excel worksheet by calculating its weight.
        /// Returns a tuple containing the worksheet name and the calculated weight.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet to process.</param>
        /// <returns>
        /// A task that, when completed, returns a tuple of the worksheet's name and its weight.
        /// </returns>
        private Task<Tuple<string, decimal>> ProcessWorksheet(Worksheet worksheet) =>
            Task.Run(() => new Tuple<string, decimal>(worksheet.Name, CalculateWorksheetWeight(worksheet)));

        /// <summary>
        /// Calculates the weight of the specified Excel worksheet.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet for which to calculate the weight.</param>
        /// <returns>
        /// A decimal value representing the worksheet's weight.
        /// </returns>
        private decimal CalculateWorksheetWeight(Worksheet worksheet)
        {
            decimal totalWeight = decimal.Zero;
            totalWeight += SumAllCellValues(worksheet) / 100;
            totalWeight += SumAllCommentsValues(worksheet) / 100;
            totalWeight += SumAllShapes(worksheet) / 100;
            totalWeight += SumAllLinks(worksheet) / 100;
            // Release COM objects, if necessary
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            return totalWeight;
        }

        /// <summary>
        /// Sums a base and optional character-based weighting for all hyperlinks in the specified worksheet.
        /// Awards 3 points for each hyperlink plus additional points for the hyperlink's address length.
        /// </summary>
        /// <param name="worksheet">The <see cref="Excel.Worksheet"/> containing hyperlinks.</param>
        /// <returns>
        /// A decimal representing the total weight of all hyperlinks.
        /// </returns>
        private static decimal SumAllLinks(Worksheet worksheet)
        {
            decimal totalWeight = decimal.Zero;
            Hyperlinks hyperlinks = worksheet.Hyperlinks;
            if (hyperlinks != null)
            {
                foreach (Hyperlink hyperlink in hyperlinks)
                {
                    // Base weighting: 3 points per hyperlink
                    totalWeight += 3;
                    // Additional weighting based on address length
                    if (!string.IsNullOrEmpty(hyperlink.Address))
                    {
                        totalWeight += hyperlink.Address.Length;
                    }
                }
            }
            return totalWeight;
        }

        /// <summary>
        /// Calculates a total weight for all shapes in the specified worksheet.
        /// Uses a fixed base of 40 points per shape.
        /// </summary>
        /// <param name="worksheet">The <see cref="Excel.Worksheet"/> containing shapes.</param>
        /// <returns>
        /// A decimal representing the cumulative weight of all shapes.
        /// </returns>
        private static decimal SumAllShapes(Worksheet worksheet)
        {
            decimal totalWeight = decimal.Zero;
            Shapes shapes = worksheet.Shapes;
            if (shapes != null)
            {
                foreach (Shape unused in shapes)
                {
                    // Example weighting: 40 points per shape
                    totalWeight += 40;
                }
            }
            return totalWeight;
        }

        /// <summary>
        /// Calculates the cumulative weight by summing the length of all comment text in the specified worksheet.
        /// </summary>
        /// <param name="worksheet">The <see cref="Excel.Worksheet"/> containing comments.</param>
        /// <returns>
        /// A decimal representing the total weight of all comment text.
        /// </returns>
        private static decimal SumAllCommentsValues(Worksheet worksheet)
        {
            decimal totalWeight = decimal.Zero;
            Comments comments = worksheet.Comments;
            if (comments != null)
            {
                foreach (Comment comment in comments)
                {
                    string commentText = comment.Text();
                    if (!string.IsNullOrEmpty(commentText))
                    {
                        totalWeight += commentText.Length;
                    }
                }
            }
            return totalWeight;
        }

        /// <summary>
        /// Iterates through all cells in the specified worksheet and sums values that can be converted to decimals.
        /// </summary>
        /// <param name="worksheet">The <see cref="Excel.Worksheet"/> to process.</param>
        /// <returns>
        /// A decimal representing the sum of all numeric cell values.
        /// </returns>
        private static decimal SumAllCellValues(Worksheet worksheet)
        {
            // Estimated weight per element
            // Average weight per cell
            const decimal weightPerCell = 1;
            return worksheet.UsedRange.Count * weightPerCell;
        }

    }
}