using System;
using System.IO;
using System.Text;
using ExcelAnalyze.Descriptor;

namespace ExcelAnalyze.Logic
{
    public abstract class ExcelFileServiceBase
    {
        protected static void AddFileSizeInfo(StringBuilder result, long fileBytes, string path)
        {
            result.AppendLine();
            var fileName = Path.GetFileName(path);
            result.AppendLine(string.Format(Message.BaseResultMessage, fileName, Helper.GetKbFromBytes(fileBytes),
                Helper.GetMbFromBytes(fileBytes), Constant.FullPercent));
        }

        protected static void AddFooterInfo(StringBuilder result)
        {
            result.AppendLine(Message.Line);
            result.AppendLine();
            result.AppendLine(string.Concat(Message.ProcessFinished, DateTime.Now));
            result.AppendLine(Message.Note);
        }

        protected static StringBuilder CreateBaseStringBuilder()
        {
            StringBuilder result = new StringBuilder();
            result.AppendLine(string.Concat(Message.ProcessStartedTime, DateTime.Now));
            return result;
        }
    }
}