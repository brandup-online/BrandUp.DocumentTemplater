using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    internal class Foreach : IDocxCommand
    {
        public string Name => "foreach";
        public HandleResult Execute(List<string> properties, object dataContext)
        {
            var items = new List<object>();

            if (dataContext is System.Collections.IEnumerable collection)
            {
                foreach (object item in collection)
                    items.Add(item);
            }

            return new(dataContext, items);
        }
    }
}
