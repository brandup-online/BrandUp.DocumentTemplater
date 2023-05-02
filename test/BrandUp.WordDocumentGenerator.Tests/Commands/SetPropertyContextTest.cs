using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    public class SetPropertyContextTest
    {
        /// <summary>
        /// Тест команды setcontextofproperty с параметром контекста "Top"
        /// </summary>
        [Fact]
        public void Success()
        {
            var data = new
            {
                Top = new
                {
                    Prop = "Раз",
                },
                Bottom = new
                {
                    Level1 = new
                    {
                        Prop = "Четыре"
                    }
                }
            };

            var command = new SetPropertyContext();
            var result = command.Execute(new() { "Top" }, data);

            Assert.Same(data.Top, result.DataContext);
            Assert.Null(result.OutputContent);
            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.None, result.OutputType);
        }
    }
}
