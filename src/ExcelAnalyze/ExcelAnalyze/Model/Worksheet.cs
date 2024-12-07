using ExcelAnalyze.Descriptor;

namespace ExcelAnalyze.Model
{
    public class Worksheet
    {
        public string Name { get; }

        public double SizeInKb { get; }

        public double SizeInMb { get; }

        public long PercentWeight { get; }

        public Worksheet(string name, long usedRange, long totalUsedRange, long fileBytes)
        {
            Name = name;
            var bytes = usedRange * fileBytes / totalUsedRange;
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