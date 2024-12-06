using System;
using System.Windows.Forms;

namespace ExcelAnalyze
{
    public static class Helper
    {
        public static void ShowError(Exception ex)
        {
            MessageBox.Show(ex.Message, Descriptor.Message.Error.Caption,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static double GetKbFromBytes(long bytes)
        {
            return bytes / 1024.0;
        }

        public static double GetMbFromKb(double kb)
        {
            return kb / 1024.0;
        }

        public static double GetMbFromBytes(long bytes)
        {
            return GetMbFromKb(GetKbFromBytes(bytes));
        }
    }
}