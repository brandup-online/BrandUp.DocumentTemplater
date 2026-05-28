using BrandUp.DocumentTemplater.Exceptions;

namespace BrandUp.DocumentTemplater
{
    internal static class TypeExtension
    {
        /// <summary>
        /// Возваращает значение свойства объекта.
        /// </summary>
        /// <param name="obj">Объект.</param>
        /// <param name="propName">Имя свойства</param>
        /// <returns>Значение свойства <c>propName</c>.</returns>
        /// <exception cref="InvalidPropertyNameException"></exception>
        /// <exception cref="ContextValueNullException"></exception>
        public static object GetPropertyValue(this object obj, string propName)
        {
            ArgumentNullException.ThrowIfNull(obj);
            ArgumentNullException.ThrowIfNull(propName);

            var type = obj.GetType();

            if (type.IsAssignableTo(typeof(IDictionary<string, object>)))
            {
                if (!((IDictionary<string, object>)obj).TryGetValue(propName, out var value))
                    throw new InvalidPropertyNameException(propName);
                return value;
            }
            else
            {
                var props = propName.Split('.');
                object result = obj;
                foreach (var prop in props)
                {
                    var property = result.GetType().GetProperty(prop) ?? throw new InvalidPropertyNameException(type, propName);
                    result = property.GetValue(result) ?? throw new ContextValueNullException();
                }
                return result;
            }
        }
    }
}