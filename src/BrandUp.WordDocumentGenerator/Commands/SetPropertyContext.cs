using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Sets property as DataContext
    /// </summary>
    internal class SetPropertyContext : IDocxCommand
    {
        public string Name => "context";
        public HandleResult Execute(List<string> properties, object dataContext)
        {
            object value = dataContext.GetType().GetValueFromContext(properties[0], dataContext);

            return new(value);
        }
    }
}
