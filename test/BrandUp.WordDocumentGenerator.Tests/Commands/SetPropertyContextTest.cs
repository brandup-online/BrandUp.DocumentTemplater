using BrandUp.DocumentTemplater.Exceptions;
using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    public class SetPropertyContextTest
    {
        [Fact]
        public void EmptyParameters_Throws()
        {
            var command = new SetPropertyContext();

            Assert.Throws<ArgumentException>(() => command.Execute(new(), new { X = 1 }));
        }

        [Fact]
        public void NullDataContext_Throws()
        {
            var command = new SetPropertyContext();

            Assert.Throws<ContextValueNullException>(() => command.Execute(new() { "X" }, null));
        }

        /// <summary>
        /// null-значение, полученное из словаря, должно превращаться
        /// в ContextValueNullException (новый guard в Execute).
        /// </summary>
        [Fact]
        public void NullPropertyValue_Throws()
        {
            var command = new SetPropertyContext();
            var data = new Dictionary<string, object> { ["X"] = null };

            Assert.Throws<ContextValueNullException>(() => command.Execute(new() { "X" }, data));
        }

        /// <summary>
        /// Тест команды <see cref="SetPropertyContext"/> с параметром контекста "Top". Контекст данных, объект.
        /// </summary>
        [Fact]
        public void Success_Object()
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

        /// <summary>
        /// Тест команды <see cref="SetPropertyContext"/> с параметром контекста "Top". Контекст данных, словарь.
        /// </summary>
        [Fact]
        public void Success_Dictionary()
        {
            var data = new Dictionary<string, object>()
            {
                { "Top" , new Dictionary<string, object>{ { "Prop", "Раз" } } },
                { "Bottom", "Четыре" }
            };

            var command = new SetPropertyContext();
            var result = command.Execute(new() { "Top" }, data);

            Assert.Same(data["Top"], result.DataContext);
            Assert.Null(result.OutputContent);
            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.None, result.OutputType);
        }
    }
}
