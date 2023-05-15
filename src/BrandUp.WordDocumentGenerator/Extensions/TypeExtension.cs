using BrandUp.DocumentTemplater.Exeptions;

namespace BrandUp.DocumentTemplater
{
    internal static class TypeExtension
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
                if (!((IDictionary<string, object>)dataContext).TryGetValue(propName, out var value))
                    throw new InvalidPropertyNameException(propName);
                return value;
            }
            else
            {
                var p = type.GetProperty(propName) ?? throw new InvalidPropertyNameException(type, propName);

                return p.GetValue(dataContext) ?? throw new ContextValueNullException();
            }
        }
    }
}
