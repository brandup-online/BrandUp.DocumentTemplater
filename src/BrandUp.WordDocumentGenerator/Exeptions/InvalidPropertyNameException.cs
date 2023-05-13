namespace BrandUp.DocumentTemplater.Exeptions
{
    public class InvalidPropertyNameException : Exception
    {
        public InvalidPropertyNameException(Type type, string propertyName)
            : base($"Не найдено свойство {propertyName} у объекта с типом {type.FullName}.")
        { }

        public InvalidPropertyNameException(string key) : base($"Не найденно значение с ключем {key}") { }
    }
}
