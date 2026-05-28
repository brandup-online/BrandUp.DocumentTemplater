namespace BrandUp.DocumentTemplater
{
    internal static class ObjectExtension
    {
        /// <summary>
        /// Приводит объект в стороку с соответствующем форматом
        /// </summary>
        /// <param name="value">Объект</param>
        /// <param name="format">Формат</param>
        /// <returns> Форматированую строку</returns>
        public static string ToString(this object value, string format)
        {
            if (string.IsNullOrEmpty(format))
                return value.ToString();

            if (value is bool boolean)
            {
                if (format == "b")
                    return boolean ? "да" : "нет";
                else if (format == "B")
                    return boolean ? "Да" : "Нет";
                else
                    return value.ToString();
            }

            // Покрывает все числовые типы (int, long, short, decimal, double, float, …),
            // а также DateTime, TimeSpan, Guid и прочие IFormattable.
            if (value is IFormattable formattable)
                return formattable.ToString(format, null);

            return string.Format(format, value);
        }
    }
}