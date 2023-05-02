using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Sets collection as DataContext
    /// </summary>
    internal class Foreach : IContextCommand
    {
        #region IContextCommand members
        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        public string Name => "foreach";

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        public HandleResult Execute(List<string> parameters, object dataContext)
        {
            var items = new List<object>();
            var value = dataContext;
            if (parameters.Any())
                value = value.GetType().GetValueFromContext(parameters[0], value) ?? dataContext;

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
