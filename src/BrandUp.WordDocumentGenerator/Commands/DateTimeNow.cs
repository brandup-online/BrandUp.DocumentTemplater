using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    /// <summary>
    /// Устанавливает текущее дату и время
    /// </summary>
    internal class DateTimeNow : ITemplaterCommand
    {
        #region IContextCommand members

        public string Name => "datetimenow";

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