using DocumentFormat.OpenXml;

namespace BrandUp.DocumentTemplater.Internals
{
    /// <summary>
    /// OpenXml element and data context
    /// </summary>
    internal class OpenXmlElementDataContext : ICloneable
    {
        public OpenXmlElement Element { get; }
        public object DataContext { get; set; }

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
