using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Sets property as DataContext
    /// </summary>
    internal class SetPropertyContext : IContextCommand
    {
        #region IContextCommand members

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        public string Name => "setcontextofproperty";
        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        public HandleResult Execute(List<string> parameters, object dataContext)
        {
            object value = dataContext.GetType().GetValueFromContext(parameters[0], dataContext);

            return new(value);
        }

        #endregion
    }
}
