namespace BrandUp.DocumentTemplater.Exceptions
{
    public class ContextValueNullException : Exception
    {
        public ContextValueNullException() : base("Значение полученое из контекста данных равно null.") { }
    }
}
