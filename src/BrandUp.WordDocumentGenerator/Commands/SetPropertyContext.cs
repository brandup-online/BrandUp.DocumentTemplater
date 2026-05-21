using BrandUp.DocumentTemplater.Abstraction;
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
            object value = dataContext.GetPropertyValue(parameters[0]);

            return new(value);
        }

        #endregion
    }
}