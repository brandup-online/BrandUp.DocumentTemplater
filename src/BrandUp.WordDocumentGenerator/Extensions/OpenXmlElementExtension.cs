using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BrandUp.DocumentTemplater
{
    internal static class OpenXmlElementExtension
    {
        /// <summary>
        /// Determines whether <see cref="OpenXmlElement"> is content control (the specified open XML element data context).
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        /// <returns>
        /// <c>true</c> if <see cref="OpenXmlElement"> is content control (the specified open XML element data context); otherwise, <c>false</c>.
        /// </returns>
        public static bool IsContentControl(this OpenXmlElement element)
        {
            if (element == null)
                return false;

            return element is SdtBlock || element is SdtRun || element is SdtRow || element is SdtCell;
        }
    }
}
