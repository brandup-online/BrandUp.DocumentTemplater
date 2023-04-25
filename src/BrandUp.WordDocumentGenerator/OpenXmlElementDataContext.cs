using DocumentFormat.OpenXml;

namespace BrandUp.DocxGenerator
{
    /// <summary>
    /// OpenXml element and data context
    /// </summary>
    public class OpenXmlElementDataContext : ICloneable
    {
        private OpenXmlElement element;
        private object dataContext = null;

        public OpenXmlElement Element
        {
            get { return this.element; }
        }

        public object DataContext
        {
            get { return this.dataContext; }
            internal set { this.dataContext = value; }
        }

        public OpenXmlElementDataContext(OpenXmlElement element, object dataContext)
        {
            if (element == null)
                throw new ArgumentNullException("element");

            this.element = element;
            this.dataContext = dataContext;
        }

        public object Clone()
        {
            OpenXmlElementDataContext ret = new OpenXmlElementDataContext(this.Element, this.DataContext);
            return ret;
        }

        public OpenXmlElementDataContext CloneTyped()
        {
            return (OpenXmlElementDataContext)this.Clone();
        }
    }
}
