namespace BrandUp.DocumentTemplater.Exeptions
{
    internal class InvalidCommandException : Exception
    {
        public InvalidCommandException(string commandName) : base($"Неизвестная команда: {commandName}") { }
    }
}
