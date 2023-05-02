namespace BrandUp.DocumentTemplater.Handling
{
    /// <summary>
    /// Результат выполнения команды
    /// </summary>
    public class HandleResult
    {
        /// <summary>
        /// Новый контекст выполнения
        /// </summary>
        public object DataContext { get; }

        /// <summary>
        /// Строка для записи в файл
        /// </summary>
        public string OutputContent { get; set; }

        /// <summary>
        /// Список данных для записи в файл
        /// </summary>
        public List<object> OutputList { get; set; }

        /// <summary>
        /// Тип изменения документа
        /// </summary>
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
