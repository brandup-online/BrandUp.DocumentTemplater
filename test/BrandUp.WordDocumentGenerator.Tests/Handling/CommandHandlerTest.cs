using BrandUp.DocumentTemplater.Abstraction;
using BrandUp.DocumentTemplater.Commands;
using BrandUp.DocumentTemplater.Exceptions;

namespace BrandUp.DocumentTemplater.Handling
{
    public class CommandHandlerTest
    {
        /// <summary>
        /// Имена команд должны разрешаться без учёта регистра
        /// (фиксирует переход словаря на OrdinalIgnoreCase).
        /// </summary>
        [Theory]
        [InlineData("prop")]
        [InlineData("PROP")]
        [InlineData("Prop")]
        public void Handle_IsCaseInsensitive(string name)
        {
            var result = CommandHandler.Handle(name, new(), 42);

            Assert.Equal("42", result.OutputContent);
        }

        [Fact]
        public void Handle_UnknownCommand_Throws()
        {
            Assert.Throws<InvalidCommandException>(() => CommandHandler.Handle("doesnotexist", new(), 42));
        }

        [Fact]
        public void Handle_NullCommandName_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => CommandHandler.Handle(null, new(), 42));
        }

        [Fact]
        public void Handle_NullProperties_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => CommandHandler.Handle("prop", null, 42));
        }

        [Fact]
        public void Handle_NullDataContext_Throws()
        {
            Assert.Throws<ContextValueNullException>(() => CommandHandler.Handle("prop", new(), null));
        }

        /// <summary>
        /// "prop" уже зарегистрирован в статическом конструкторе.
        /// </summary>
        [Fact]
        public void AddHandler_DuplicateName_Throws()
        {
            Assert.Throws<ArgumentException>(() => CommandHandler.AddHandler(new Prop()));
        }

        /// <summary>
        /// Регистрация дубликата должна срабатывать и при отличающемся регистре.
        /// </summary>
        [Fact]
        public void AddHandler_DuplicateName_IsCaseInsensitive()
        {
            Assert.Throws<ArgumentException>(() => CommandHandler.AddHandler(new UpperCasePropStub()));
        }

        class UpperCasePropStub : ITemplaterCommand
        {
            public string Name => "PROP";
            public HandleResult Execute(List<string> parameters, object dataContext) => new(dataContext);
        }
    }
}
