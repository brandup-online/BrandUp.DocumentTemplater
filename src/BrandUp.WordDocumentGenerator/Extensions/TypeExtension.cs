namespace BrandUp.DocumentTemplater
{
    public static class TypeExtension
    {
        /// <summary>
        /// Возваращает данные из контектста
        /// </summary>
        /// <param name="type">Тип данных</param>
        /// <param name="propName">Имя свойства</param>
        /// <param name="dataContext">Контекст данных</param>
        /// <returns> Значение свойства <c>propName</c></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public static object GetValueFromContext(this Type type, string propName, object dataContext)
        {
            if (type.IsAssignableTo(typeof(IDictionary<string, object>)))
            {
                return ((IDictionary<string, object>)dataContext)[propName];
            }
            else
            {
                try
                {
                    var p = type.GetProperty(propName) ?? throw new InvalidOperationException("Не найдено свойство " + propName + " у обхекта с типом " + type.FullName + ".");

                    return p.GetValue(dataContext) ?? string.Empty;
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }
    }
}
