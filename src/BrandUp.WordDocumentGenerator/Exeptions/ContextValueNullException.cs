namespace BrandUp.DocumentTemplater.Exeptions
{
    public class ContextValueNullException : Exception
    {
        public ContextValueNullException() : base("Значение полученое из контекста данных равно null.") { }
    }
}
