using BrandUp.DocumentTemplater.Handling;

namespace BrandUp.DocumentTemplater.Commands
{
    public class PropTest
    {
        private static readonly List<object[]> contexts = new()
        {
            new object[] { "String" },
            new object[] { new DateTime(2023, 12, 22, 13, 55, 16) },
            new object[] { new TimeSpan(13, 55, 16) },
            new object[] { 12 },
            new object[] { 123123123123m },
            new object[] { 0.9200f },
            new object[] { double.MaxValue },
            new object[] { true }
        };

        /// <summary>
        /// Значения всех поддерживаемых объектов. 
        /// </summary>
        public static IEnumerable<object[]> TestContexts => contexts;

        /// <summary>
        /// Выполнение команды <see cref="Prop"/>, без параметров.
        /// </summary>
        /// <param name="context">Контекст данных</param>
        [Theory]
        [MemberData(nameof(TestContexts))]
        public void SuccessPatamtless(object context)
        {
            var command = new Prop();

            var result = command.Execute(new() { }, context);

            Assert.Equal(context, result.DataContext);
            Assert.Equal(context.ToString(), result.OutputContent);
            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.Content, result.OutputType);
        }

        /// <summary>
        /// Выполнение команды <see cref="Prop"/> с одним параметром - имя свойства. Дата контекст - словарь
        /// </summary>
        /// <param name="key">Имя свойства</param>
        /// <param name="format">Формат</param>
        /// <param name="value">Результат</param>
        [Theory]
        [InlineData("FullDate")]
        [InlineData("Span")]
        [InlineData("IntValue")]
        [InlineData("DecValue")]
        [InlineData("FloatValue")]
        [InlineData("DoubleValue")]
        [InlineData("Yes")]
        public void SuccessKey(string key)
        {
            var data = new Dictionary<string, object>
              {
                  { "FullDate" , new DateTime(2023,12, 22, 13, 55, 16) },
                  { "Span" , new TimeSpan(13, 55, 16) },
                  { "IntValue" , 12 },
                  { "DecValue" ,123123123123m },
                  { "FloatValue" , 0.9200f },
                  { "DoubleValue" , double.MaxValue },
                  { "Yes" , true }
              };
            var command = new Prop();

            var result = command.Execute(new() { key }, data);

            Assert.Equal(data[key], result.DataContext);
            Assert.Equal(data[key].ToString(), result.OutputContent);
            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.Content, result.OutputType);
        }

        /// <summary>
        /// Выполнение команды <see cref="Prop"/> с двумя парамтрами: имя свойства, формат. Дата контекст словарь
        /// </summary>
        /// <param name="key">Имя свойства</param>
        /// <param name="format">Формат</param>
        /// <param name="value">Результат</param>
        [Theory]
        [InlineData("FullDate", "dd-MM-yyyy hh-mm-ss", "22-12-2023 01-55-16")]
        [InlineData("Span", "hh\\:mm\\.ss", "13:55.16")]
        [InlineData("IntValue", "# градусов", "12 градусов")]
        [InlineData("DecValue", "\\##\\#", "#123123123123#")]
        [InlineData("FloatValue", "#.##", ".92")]
        [InlineData("DoubleValue", "0.0e+00", "1.8e+308")]
        [InlineData("Yes", "b", "да")]
        [InlineData("No", "B", "Нет")]
        public void SuccessDictionaryFormats(string key, string format, string value)
        {
            Thread.CurrentThread.CurrentCulture = new("en-US");

            var data = new Dictionary<string, object>
            {
                { "FullDate" , new DateTime(2023,12, 22, 13, 55, 16) },
                { "Span" , new TimeSpan(13, 55, 16) },
                { "IntValue" , 12 },
                { "DecValue" ,123123123123m },
                { "FloatValue" , 0.9200f },
                { "DoubleValue" , double.MaxValue },
                { "Yes" , true },
                { "No" , false },
            };

            var command = new Prop();

            var result = command.Execute(new() { key, format }, data);

            Assert.Same(data[key], result.DataContext);
            Assert.Equal(value, result.OutputContent);
            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.Content, result.OutputType);
        }

        /// <summary>
        /// Выполнение команды <see cref="Prop"/> с двумя парамтрами: имя свойства, формат. Дата контекст объект.
        /// </summary>
        /// <param name="key">Имя свойства</param>
        /// <param name="format">Формат</param>
        /// <param name="value">Результат</param>
        [Theory]
        [InlineData("FullDate", "dd-MM-yyyy hh-mm-ss", "22-12-2023 01-55-16")]
        [InlineData("Span", "hh\\:mm\\.ss", "13:55.16")]
        [InlineData("IntValue", "# градусов", "12 градусов")]
        [InlineData("DecValue", "\\##\\#", "#123123123123#")]
        [InlineData("FloatValue", "#.##", ".92")]
        [InlineData("DoubleValue", "0.0e+00", "1.8e+308")]
        [InlineData("Yes", "b", "да")]
        [InlineData("No", "B", "Нет")]
        public void SuccessObjectFormats(string key, string format, string value)
        {
            Thread.CurrentThread.CurrentCulture = new("en-US");

            var data = new
            {
                FullDate = new DateTime(2023, 12, 22, 13, 55, 16),
                Span = new TimeSpan(13, 55, 16),
                IntValue = 12,
                DecValue = 123123123123m,
                FloatValue = 0.9200f,
                DoubleValue = double.MaxValue,
                Yes = true,
                No = false
            };

            var command = new Prop();

            var result = command.Execute(new() { key, format }, data);

            Assert.Equal(data.GetType().GetProperty(key).GetValue(data), result.DataContext);
            Assert.Equal(value, result.OutputContent);
            Assert.Null(result.OutputList);
            Assert.Equal(CommandOutputType.Content, result.OutputType);
        }
    }
}
