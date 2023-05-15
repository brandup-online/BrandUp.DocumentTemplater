using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Commands;
using BrandUp.DocumentTemplater.Exeptions;

namespace BrandUp.DocumentTemplater.Handling
{
    internal static class CommandHandler
    {
        readonly static IDictionary<string, ITemplaterCommand> commands = new Dictionary<string, ITemplaterCommand>();

        static CommandHandler()
        {
            AddHandler(new SetPropertyContext());
            AddHandler(new Foreach());
            AddHandler(new Prop());
            AddHandler(new DateTimeNow());
        }

        /// <summary>
        /// Добавляет обработчик команды
        /// </summary>
        /// <param name="command"></param>
        /// <exception cref="ArgumentException"></exception>
        public static void AddHandler(ITemplaterCommand command)
        {
            if (!commands.TryAdd(command.Name.ToLower(), command))
            {
                throw new ArgumentException("Handler with this name already exist.", nameof(command));
            }
        }

        /// <summary>
        /// Обрабатывает команду
        /// </summary>
        /// <param name="commandName">Название команды из шаблона</param>
        /// <param name="properties">Параметры команды</param>
        /// <param name="dataContext">Контекст данных</param>
        /// <returns><see cref="HandleResult"/></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static HandleResult Handle(string commandName, List<string> properties, object dataContext)
        {
            if (commandName == null)
                throw new ArgumentNullException(nameof(commandName));
            if (properties == null)
                throw new ArgumentNullException(nameof(properties));
            if (dataContext == null)
                throw new ContextValueNullException();

            if (commands.TryGetValue(commandName.ToLower(), out var command))
                return command.Execute(properties, dataContext);
            else throw new InvalidCommandException(commandName);
        }
    }
}