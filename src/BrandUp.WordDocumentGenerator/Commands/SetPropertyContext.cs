using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Exceptions;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Задает свойство как контектст данных
    /// </summary>
    internal class SetPropertyContext : ITemplaterCommand
    {
        #region ITemplaterCommand members

        public string Name => "context";
        public HandleResult Execute(List<string> parameters, object dataContext)
        {
            if (dataContext == null)
                throw new ContextValueNullException();
            if (parameters.Count == 0)
                throw new ArgumentException("Команда 'context' требует имя свойства.", nameof(parameters));

            object value = dataContext.GetPropertyValue(parameters[0]) ?? throw new ContextValueNullException();

            return new(value);
        }

        #endregion
    }
}