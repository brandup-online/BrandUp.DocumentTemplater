using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Exeptions;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Устанавливает коллекцию как контекст данных
    /// </summary>
    internal class Foreach : ITemplaterCommand
    {
        #region IContextCommand members

        public string Name => "foreach";

        public HandleResult Execute(List<string> parameters, object dataContext)
        {
            var items = new List<object>();
            var value = dataContext ?? throw new ContextValueNullException();
            if (parameters.Count > 0)
                value = value.GetType().GetValueFromContext(parameters[0], dataContext) ?? throw new ContextValueNullException();

            if (value is System.Collections.IEnumerable collection)
            {
                foreach (object item in collection)
                    items.Add(item);
            }

            return new(dataContext, items);
        }

        #endregion
    }
}
