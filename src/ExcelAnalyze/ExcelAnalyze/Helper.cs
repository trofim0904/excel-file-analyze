using System;
using System.Windows.Forms;

namespace ExcelAnalyze
{
    /// <summary>
    /// Provides helper methods for error display and unit conversion (bytes to KB/MB).
    /// </summary>
    public static class Helper
    {
        /// <summary>
        /// Displays an error message in a message box.
        /// </summary>
        /// <param name="ex">The exception containing the error message.</param>
        public static void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message, Descriptor.Message.Error.Caption,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Converts the specified byte size to kilobytes.
        /// </summary>
        /// <param name="bytes">The number of bytes to convert.</param>
        /// <returns>
        /// The equivalent size in kilobytes.
        /// </returns>
        public static double GetKbFromBytes(long bytes)
        {
            return bytes / 1024.0;
        }

        /// <summary>
        /// Converts the specified kilobyte size to megabytes.
        /// </summary>
        /// <param name="kb">The size in kilobytes.</param>
        /// <returns>
        /// The equivalent size in megabytes.
        /// </returns>
        public static double GetMbFromKb(double kb)
        {
            return kb / 1024.0;
        }

        /// <summary>
        /// Converts the specified byte size to megabytes.
        /// </summary>
        /// <param name="bytes">The number of bytes to convert.</param>
        /// <returns>
        /// The equivalent size in megabytes.
        /// </returns>
        public static double GetMbFromBytes(long bytes)
        {
            return GetMbFromKb(GetKbFromBytes(bytes));
        }
    }
}
