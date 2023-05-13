using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Sets value to palceholder
    /// </summary>
    internal class Prop : IDocxCommand
    {
        public string Name => "prop";
        public HandleResult Execute(List<string> properties, object dataContext)
        {
            string output = null;
            object value;
            if (properties.Count > 0)
                value = dataContext.GetType().GetValueFromContext(properties[0], dataContext);
            else value = dataContext;

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
    }
}
