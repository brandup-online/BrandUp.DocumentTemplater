using DocumentFormat.OpenXml;

namespace BrandUp.DocumentTemplater.Internals
{
    /// <summary>
    /// OpenXml элемент и связанный с ним контекст данных
    /// </summary>
    internal class OpenXmlElementDataContext : ICloneable
    {
        /// <summary>
        /// OpenXml элемент 
        /// </summary>
        public OpenXmlElement Element { get; }
        /// <summary>
        /// Контекст данных
        /// </summary>
        public object DataContext { get; }

        public OpenXmlElementDataContext(OpenXmlElement element)
        {
            Element = element ?? throw new ArgumentNullException(nameof(element));
        }

        public OpenXmlElementDataContext(OpenXmlElement element, object dataContext) : this(element)
        {
            DataContext = dataContext ?? throw new ArgumentNullException(nameof(dataContext));
        }

        public OpenXmlElementDataContext CloneTyped()
        {
            return (OpenXmlElementDataContext)Clone();
        }

        #region ICloneable members

        public object Clone()
        {
            OpenXmlElementDataContext ret = new(Element, DataContext);
            return ret;
        }

        #endregion
    }
}
