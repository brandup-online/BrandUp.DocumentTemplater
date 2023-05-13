using BrandUp.DocumentTemplater.Exeptions;
using System.Collections;

namespace BrandUp.DocumentTemplater
{
    public static class TypeExtension
    {
        /// <summary>
        /// Gets property value from data.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="propName">name of property</param>
        /// <param name="dataContext">Model with data</param>
        /// <returns> Value by <c>propName</c></returns>
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
                try
                {
                    var p = type.GetProperty(propName) ?? throw new InvalidPropertyNameException(type, propName);

                    return p.GetValue(dataContext) ?? string.Empty;
                }
                catch (Exception)
                {
                    return string.Empty;
                }
            }
        }

        public static bool IsContentType(this Type type)
        {
            return type.IsValueType || type == typeof(string) || type.IsAssignableTo(typeof(IEnumerable));
        }
    }
}
