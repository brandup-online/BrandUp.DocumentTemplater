namespace BrandUp.DocumentTemplater.Exeptions
{
    public class DataContextNullException : Exception
    {
        public DataContextNullException() : base("Контекст данных равен null.") { }
    }
}
