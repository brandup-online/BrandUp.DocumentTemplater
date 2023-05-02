using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Устанавливает текущее дату и время
    /// </summary>
    internal class DateTimeNow : IContextCommand
    {
        #region IContextCommand members

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        public string Name => "datetimenow";

        /// <summary>
        /// <inheritdoc/>
        /// </summary>
        public HandleResult Execute(List<string> parameters, object dataContext)
        {
            string output;
            DateTime d = DateTime.Now;

            if (parameters.Count > 0)
                output = d.ToString(parameters[0]);
            else
                output = d.ToString();

            return new(dataContext, output);
        }

        #endregion
    }
}
