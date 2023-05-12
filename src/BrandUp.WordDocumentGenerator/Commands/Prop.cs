using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    ///  Устанивливает значение в соответствующий элемент управлнеие
    /// </summary>
    internal class Prop : ITemplaterCommand
    {
        #region IContextCommand members

        public string Name => "prop";

        public HandleResult Execute(List<string> parameters, object dataContext)
        {
            string output = null;
            object value;
            if (parameters.Any())
                value = dataContext.GetType().GetValueFromContext(parameters[0], dataContext);
            else value = dataContext;

            if (value != null)
            {
                if (parameters.Count > 1 && !string.IsNullOrEmpty(parameters[1]))
                {
                    var format = parameters[1];
                    output = value.ToString(format);
                }
                else
                    output = value.ToString();
            }

            return new(value, output);
        }

        #endregion
    }
}
