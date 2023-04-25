namespace BrandUp.DocxGenerator
{
    public class FoundCommandEventArgs
    {
        private object dataContext = null;
        private bool isChangedDataContext = false;

        public string CommandName
        {
            get;
            private set;
        }

        public List<string> CommandArguments
        {
            get;
            private set;
        }

        public object DataContext
        {
            get { return this.dataContext; }
            internal set
            {

                this.dataContext = value;
                this.isChangedDataContext = true;
            }
        }

        public bool IsChangedDataContext
        {
            get { return this.isChangedDataContext; }
        }

        public string OutputContent
        {
            get;
            set;
        }

        public List<object> OutputList
        {
            get;
            set;
        }

        public DocumentGeneratorCommandOutputType OutputType
        {
            get;
            set;
        }

        public FoundCommandEventArgs() { }
        public FoundCommandEventArgs(string commandName, List<string> commandArguments, object dataContext)
        {
            if (string.IsNullOrEmpty(commandName))
                throw new ArgumentNullException("commandName");
            if (commandArguments == null)
                throw new ArgumentNullException("commandArguments");
            if (dataContext == null)
                throw new ArgumentNullException("dataContext");

            this.CommandName = commandName;
            this.CommandArguments = commandArguments;
            this.dataContext = dataContext;
            this.OutputType = DocumentGeneratorCommandOutputType.Content;
        }
    }
}
