using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Commands;
using BrandUp.DocumentTemplater.Exceptions;

namespace BrandUp.DocumentTemplater.Handling
{
    internal static class CommandHandler
    {
        readonly static IDictionary<string, ITemplaterCommand> commands = new Dictionary<string, ITemplaterCommand>(StringComparer.OrdinalIgnoreCase);

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
            if (!commands.TryAdd(command.Name, command))
                throw new ArgumentException("Handler with this name already exist.", nameof(command));
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
            ArgumentNullException.ThrowIfNull(commandName);
            ArgumentNullException.ThrowIfNull(properties);
            if (dataContext == null)
                throw new ContextValueNullException();

            if (commands.TryGetValue(commandName, out var command))
                return command.Execute(properties, dataContext);
            else
                throw new InvalidCommandException(commandName);
        }
    }
}