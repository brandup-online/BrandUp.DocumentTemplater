using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    internal class DateTimeNow : IDocxCommand
    {
        public string Name => "datetimenow";

        public HandleResult Execute(List<string> properties, object dataContext)
        {
            string output;
            DateTime d = DateTime.Now;

            if (properties.Count > 0)
                output = d.ToString(properties[0]);
            else
                output = d.ToString();

            return new(dataContext, output);
        }
    }
}
