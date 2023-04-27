using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Abstraction
{
    public interface IDocxCommand
    {
        public string Name { get; }
        public HandleResult Execute(List<string> properties, object dataContext);
    }
}
