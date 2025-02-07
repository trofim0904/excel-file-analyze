using System;
using System.IO;
using System.Text;
using ExcelAnalyze.Descriptor;

namespace ExcelAnalyze.Logic
{
    /// <summary>
    /// Provides base functionality for services that process Excel files, 
    /// including methods for building result messages and handling file size information.
    /// </summary>
    public abstract class ExcelFileServiceBase
    {
        /// <summary>
        /// Appends file size information to the specified <see cref="StringBuilder"/>.
        /// </summary>
        /// <param name="result">The <see cref="StringBuilder"/> to which file size info is appended.</param>
        /// <param name="fileBytes">The size of the file in bytes.</param>
        /// <param name="path">The full file path.</param>
        protected static void AddFileSizeInfo(StringBuilder result, long fileBytes, string path)
        {
            result.AppendLine();
            var fileName = Path.GetFileName(path);
            result.AppendLine(string.Format(
                Message.BaseResultMessage,
                fileName,
                Helper.GetKbFromBytes(fileBytes),
                Helper.GetMbFromBytes(fileBytes),
                Constant.FullPercent));
        }

        /// <summary>
        /// Appends footer information to the specified <see cref="StringBuilder"/>, 
        /// indicating the process has finished and providing any additional notes.
        /// </summary>
        /// <param name="result">
        /// The <see cref="StringBuilder"/> that receives the footer information.
        /// </param>
        protected static void AddFooterInfo(StringBuilder result)
        {
            result.AppendLine(Message.Line);
            result.AppendLine();
            result.AppendLine(string.Concat(Message.ProcessFinished, DateTime.Now));
            result.AppendLine(Message.Note);
        }

        /// <summary>
        /// Creates and initializes a <see cref="StringBuilder"/> with 
        /// a header message that includes the current start time.
        /// </summary>
        /// <returns>A new <see cref="StringBuilder"/> instance containing a start-time header.</returns>
        protected static StringBuilder CreateBaseStringBuilder()
        {
            StringBuilder result = new StringBuilder();
            result.AppendLine(string.Concat(Message.ProcessStartedTime, DateTime.Now));
            return result;
        }
    }

}