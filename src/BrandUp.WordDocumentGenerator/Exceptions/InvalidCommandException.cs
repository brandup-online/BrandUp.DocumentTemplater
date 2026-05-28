namespace BrandUp.DocumentTemplater.Exceptions
{
    internal class InvalidCommandException : Exception
    {
        public InvalidCommandException(string commandName) : base($"Неизвестная команда: {commandName}") { }
    }
}
