using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    public class DateTimeNowTest : TestBase
    {
        [Fact]
        public void Success()
        {
            var command = new DateTimeNow();
            int context = 2;

            var result = command.Execute(new() { }, context);

            Assert.Equal(context, result.DataContext);
            Assert.NotNull(result.OutputContent);
            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.Content, result.OutputType);
        }
    }
}
