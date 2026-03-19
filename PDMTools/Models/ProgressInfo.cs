namespace PDMTools.Models
{
    public sealed class ProgressInfo
    {
        public int Percentage { get; }
        public string Message { get; }

        public ProgressInfo(int percentage, string message)
        {
            Percentage = percentage;
            Message = message;
        }
    }
}
