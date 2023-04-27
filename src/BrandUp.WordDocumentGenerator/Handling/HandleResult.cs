namespace BrandUp.DocumentTemplater.Handling
{
    public class HandleResult
    {
        public object DataContext { get; }

        public string OutputContent { get; set; }

        public List<object> OutputList { get; set; }

        public CommandOutputType OutputType { get; set; }

        public HandleResult(object dataContext)
        {
            DataContext = dataContext ?? throw new ArgumentNullException(nameof(dataContext));
            OutputType = CommandOutputType.None;
        }

        public HandleResult(object dataContext, string outputContent) : this(dataContext)
        {
            OutputContent = outputContent ?? throw new ArgumentNullException(nameof(outputContent));
            OutputType = CommandOutputType.Content;
        }
        public HandleResult(object dataContext, List<object> outputList) : this(dataContext)
        {
            OutputList = outputList ?? throw new ArgumentNullException(nameof(outputList));
            OutputType = CommandOutputType.List;
        }
    }
}
