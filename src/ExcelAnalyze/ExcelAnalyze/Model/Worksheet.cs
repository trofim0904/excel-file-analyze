using System;
using ExcelAnalyze.Descriptor;

namespace ExcelAnalyze.Model
{
    public class Worksheet
    {
        public string Name { get; }

        public double SizeInKb { get; }

        public double SizeInMb { get; }

        public long PercentWeight { get; }

        public Worksheet(Tuple<string, decimal> excelWorksheetTuple, decimal totalWeight, long fileBytes) 
            : this(excelWorksheetTuple.Item1, excelWorksheetTuple.Item2, totalWeight, fileBytes) { }

        public Worksheet(string name, decimal weight, decimal totalWeight, long fileBytes)
        {
            Name = name;
            var bytes = (long) (weight * fileBytes / totalWeight);
            SizeInKb = Helper.GetKbFromBytes(bytes);
            SizeInMb = Helper.GetMbFromBytes(bytes);
            PercentWeight = bytes * Constant.FullPercent / fileBytes;
        }

        public override string ToString()
        {
            return string.Format(Message.BaseResultMessage, Name, SizeInKb, SizeInMb, PercentWeight);
        }

    }
}