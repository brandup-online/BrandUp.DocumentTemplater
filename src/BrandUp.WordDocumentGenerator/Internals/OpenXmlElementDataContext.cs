using BrandUp.DocumentTemplater.Exeptions;
using DocumentFormat.OpenXml;

namespace BrandUp.DocumentTemplater.Internals
{
    /// <summary>
    /// OpenXml элемент и связанный с ним контекст данных
    /// </summary>
    internal class OpenXmlElementDataContext(OpenXmlElement element) : ICloneable
    {
        /// <summary>
        /// OpenXml элемент 
        /// </summary>
        public OpenXmlElement Element { get; } = element ?? throw new ArgumentNullException(nameof(element));
        /// <summary>
        /// Контекст данных
        /// </summary>
        public object DataContext { get; }

        public OpenXmlElementDataContext(OpenXmlElement element, object dataContext) : this(element)
        {
            DataContext = dataContext ?? throw new ContextValueNullException();
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