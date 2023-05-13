using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Sets collection as DataContext
    /// </summary>
    internal class Foreach : IDocxCommand
    {
        public string Name => "foreach";
        public HandleResult Execute(List<string> properties, object dataContext)
        {
            var items = new List<object>();
            var value = dataContext;
            if (properties.Count > 0)
                value = value.GetType().GetValueFromContext(properties[0], value) ?? dataContext;

            if (value is System.Collections.IEnumerable collection)
            {
                foreach (object item in collection)
                    items.Add(item);
            }

            return new(dataContext, items);
        }
    }
}
