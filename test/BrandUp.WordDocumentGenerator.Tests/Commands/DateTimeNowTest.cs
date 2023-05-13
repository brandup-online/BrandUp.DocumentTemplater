using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    public class DateTimeNowTest
    {
        /// <summary>
        /// Тест комманды <see cref="DateTimeNow"/> без параметров
        /// </summary>
        [Fact]
        public void Success()
        {
            var command = new DateTimeNow();
            int context = 2;

            var result = command.Execute(new() { }, context);
            var now = DateTime.Now.ToString();

            Assert.Equal(context, result.DataContext);

            Assert.NotNull(result.OutputContent);
            Assert.Equal(result.OutputContent, now);

            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.Content, result.OutputType);
        }

        /// <summary>
        /// Тест комманды <see cref="DateTimeNow"/> c параметром "format"
        /// </summary>
        [Fact]
        public void Success_Format()
        {
            var command = new DateTimeNow();
            int context = 2;
            var format = "dd-MM-yyyy hh:mm:ss";

            var result = command.Execute(new() { format }, context);
            var now = DateTime.Now.ToString(format);

            Assert.Equal(context, result.DataContext);

            Assert.NotNull(result.OutputContent);
            Assert.Equal(result.OutputContent, now);

            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.Content, result.OutputType);
        }
    }
}
