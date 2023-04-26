namespace BrandUp.DocxGenerator.Handling
{
    internal static class CommandHandler
    {
        readonly static IDictionary<string, Func<List<string>, object, HandleResult>> commands = new Dictionary<string, Func<List<string>, object, HandleResult>>
        {
            { "setcontextofproperty", SetPropertyContext },
            { "foreach", Foreach },
            { "prop", PropValue },
            { "datetimenow",  DateTimeNow }
        };

        public static HandleResult Handle(string commandName, List<string> properties, object dataContext)
        {
            if (commandName == null)
                throw new ArgumentNullException(nameof(commandName));
            if (properties == null)
                throw new ArgumentNullException(nameof(properties));
            if (dataContext == null)
                throw new ArgumentNullException(nameof(dataContext));

            if (commands.TryGetValue(commandName.ToLower(), out var Command))
                return Command(properties, dataContext);
            else throw new ArgumentException("Wrong command", nameof(commandName));
        }

        #region Commands

        static HandleResult SetPropertyContext(List<string> properties, object dataContext)
        {
            object value = dataContext.GetType().GetValueFromContext(properties[0], dataContext);

            return new(value);
        }

        static HandleResult Foreach(List<string> properties, object dataContext)
        {
            var items = new List<object>();

            if (dataContext is System.Collections.IEnumerable collection)
            {
                foreach (object item in collection)
                    items.Add(item);
            }

            return new(dataContext, items);
        }

        static HandleResult DateTimeNow(List<string> properties, object dataContext)
        {
            string output;
            DateTime d = DateTime.Now;

            if (properties.Count > 0)
                output = d.ToString(properties[0]);
            else
                output = d.ToString();

            return new(dataContext, output);
        }

        static HandleResult PropValue(List<string> properties, object dataContext)
        {
            string output = null;
            object value = dataContext.GetType().GetValueFromContext(properties[0], dataContext);

            if (value != null)
            {
                if (properties.Count > 1 && !string.IsNullOrEmpty(properties[1]))
                {
                    var format = properties[1];
                    output = value.ToString(format);
                }
                else
                    output = value.ToString();
            }

            return new(value, output);
        }

        #endregion
    }
}
