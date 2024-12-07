namespace ExcelAnalyze.Descriptor
{
    public static class Message
    {
        public const string ProcessFinished = "The process is finished at ";
        public const string ProcessStartedTime = "The process is started at ";
        public const string Note = "Note: The size of each worksheet is approximate.";
        public const string ProcessStarted = "Process was started. Wait please...";
        public const string Line = "-------------------------------------------------------------------";
        public const string BaseResultMessage = "{0,-25}| {1,12:F2} KB | {2,10:F2} MB | {3,6}% ";

        public static class Error
        {
            public const string Caption = "Error";
        }
    }
}