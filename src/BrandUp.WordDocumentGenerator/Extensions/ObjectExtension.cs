namespace BrandUp.DocxGenerator
{
    internal static class ObjectExtension
    {
        public static string ToString(this object value, string format)
        {
            if (value is DateTime time)
                return time.ToString(format);
            else if (value is TimeSpan span)
                return span.ToString(format);
            else if (value is decimal @decimal)
                return @decimal.ToString(format);
            else if (value is double @double)
                return @double.ToString(format);
            else if (value is float single)
                return single.ToString(format);
            else if (value is int @int)
                return @int.ToString(format);
            else if (value is bool boolean)
            {
                if (format == "b")
                    return boolean ? "да" : "нет";
                else if (format == "B")
                    return boolean ? "Да" : "Нет";
                else
                    return value.ToString();
            }
            else
                return string.Format(format, value);
        }
    }
}
