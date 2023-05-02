using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BrandUp.DocumentTemplater
{
    internal static class OpenXmlElementExtension
    {
        /// <summary>
        /// Определяет что <see cref="OpenXmlElement"> является элементом управления содержимым.
        /// </summary>
        /// <param name="element">The open XML element</param>
        /// <returns>
        /// <c>true</c> если<see cref="OpenXmlElement"> является элементом управления содержимым; иначе, <c>false</c>.
        /// </returns>
        public static bool IsContentControl(this OpenXmlElement element)
        {
            if (element == null)
                return false;

            return element is SdtBlock || element is SdtRun || element is SdtRow || element is SdtCell;
        }
    }
}
