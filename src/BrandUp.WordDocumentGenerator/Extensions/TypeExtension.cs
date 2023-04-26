namespace BrandUp.DocxGenerator
{
    public static class TypeExtension
    {
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
