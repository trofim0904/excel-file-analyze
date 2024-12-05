using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Message = ExcelAnalyze.Descriptor.Message;

namespace ExcelAnalyze.Logic
{
    public class ExcelFileServiceString : IExcelFileService
    {
        public string GetWorkBookSizes(string path, string password)
        {
            try
            {
                // Initialize Excel application
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = null;
                try
                {
                    workbook = string.IsNullOrEmpty(password)
                        ? excelApp.Workbooks.Open(path, ReadOnly: true)
                        : excelApp.Workbooks.Open(path, Password: password, ReadOnly: true);
                    StringBuilder result = new StringBuilder("Worksheet Sizes:\n");
                    foreach (Excel.Worksheet worksheet in workbook.Sheets)
                    {
                        Excel.Range usedRange = worksheet.UsedRange;
                        long totalBytes = 0;
                        // Iterate through each cell in the used range
                        foreach (Excel.Range cell in usedRange)
                        {
                            // Calculate data size
                            var cellValue = cell.Value2;
                            if (cellValue != null)
                            {
                                totalBytes += Encoding.UTF8.GetByteCount(cellValue.ToString());
                            }
                            // Add size for cell formatting
                            totalBytes += CalculateCellFormattingSize(cell);
                        }
                        double sizeInKb = totalBytes / 1024.0;
                        double sizeInMb = sizeInKb / 1024.0;
                        result.AppendLine($"{worksheet.Name}: {sizeInMb:F2} MB ({sizeInKb:F2} KB)");
                        // Release COM object for current worksheet
                        Marshal.ReleaseComObject(worksheet);
                    }
                    result.AppendLine();
                    result.AppendLine(string.Concat(Message.ProcessFinished, DateTime.Now));
                    result.AppendLine(Message.Note);
                    // Display the result in a TextBox or MessageBox
                    return result.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, Message.Error.Caption, MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    // Close the workbook and quit Excel
                    workbook?.Close(false);
                    excelApp.Quit();
                    if (workbook != null)
                    {
                        Marshal.ReleaseComObject(workbook);
                    }
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Message.Error.Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return string.Concat(Message.ProcessFinished, DateTime.Now);
        }

        private long CalculateCellFormattingSize(Excel.Range cell)
        {
            // Estimate formatting size based on common formatting attributes
            long formattingSize = 0;
            if (cell.Font != null)
            {
                formattingSize += 20; // Approx size for font attributes
            }
            if (cell.Interior != null)
            {
                formattingSize += 10; // Approx size for background color
            }
            if (cell.Borders != null)
            {
                formattingSize += 15; // Approx size for borders
            }
            return formattingSize;
        }
    }
}