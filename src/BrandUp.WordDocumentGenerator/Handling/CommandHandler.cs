using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Commands;

namespace BrandUp.DocumentTemplater.Handling
{
    internal static class CommandHandler
    {
        readonly static IDictionary<string, IDocxCommand> commands = new Dictionary<string, IDocxCommand>();

        static CommandHandler()
        {
            AddHandler(new SetPropertyContext());
            AddHandler(new Foreach());
            AddHandler(new Prop());
            AddHandler(new DateTimeNow());
        }

        public static void AddHandler(IDocxCommand command)
        {
            if (!commands.TryAdd(command.Name.ToLower(), command))
            {
                throw new ArgumentException("Handler with this name already exist.", nameof(command));
            }
        }

        public static HandleResult Handle(string commandName, List<string> properties, object dataContext)
        {
            if (commandName == null)
                throw new ArgumentNullException(nameof(commandName));
            if (properties == null)
                throw new ArgumentNullException(nameof(properties));
            if (dataContext == null)
                throw new ArgumentNullException(nameof(dataContext));

            if (commands.TryGetValue(commandName.ToLower(), out var command))
                return command.Execute(properties, dataContext);
            else throw new ArgumentException("Wrong command", nameof(commandName));
        }
    }
}
