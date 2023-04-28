using BrandUp.DocumentTemplater.Handling;
using BrandUp.DocumentTemplater.Internals;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace BrandUp.DocumentTemplater
{
    public static class WordDocumentTemplater
    {
        readonly static Regex command = new(@"\{(?<command>\w+)\((?<params>.*)\)\}", RegexOptions.IgnoreCase);

        internal static IDictionary<string, List<string>> testPropertyValues;

        /// <summary>
        /// File processing .docx from the template by filling in the fields with data from the model.
        /// </summary>
        /// <param name="dataContext">Object with data</param>
        /// <param name="templateStream">Template</param>
        /// <param name="cancellationToken">Cancellation Token</param>
        /// <returns>The .docx file stream </returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<Stream> GenerateDocument(object dataContext, Stream templateStream, CancellationToken cancellationToken)
        {
            if (dataContext == null)
                throw new ArgumentNullException(nameof(dataContext));
            if (templateStream == null)
                throw new ArgumentNullException(nameof(templateStream));

            testPropertyValues = new Dictionary<string, List<string>>();

            // WordprocessingDocument changing incoming stream, so we copy data to output stream and changing it. 
            var output = new MemoryStream();
            await templateStream.CopyToAsync(output, cancellationToken);

            using (var wordDocument = WordprocessingDocument.Open(output, true))
            {
                await Task.Run(() =>
                {
                    wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                    MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
                    Document document = mainDocumentPart.Document;

                    foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                    {
                        ProcessPlaceholder(new(part.Header, dataContext));
                        part.Header.Save();
                    }

                    foreach (FooterPart part in mainDocumentPart.FooterParts)
                    {
                        ProcessPlaceholder(new(part.Footer, dataContext));
                        part.Footer.Save();
                    }

                    ProcessPlaceholder(new(document, dataContext));
                    OpenXmlHelper.EnsureUniqueContentControlIdsForMainDocumentPart(mainDocumentPart);
                    document.Save();
                }, cancellationToken);
            }

            output.Seek(0, SeekOrigin.Begin);
            return output;
        }

        /// <summary>
        /// Processes a placeholder node.
        /// </summary>
        /// <param name="openXmlElementDataContext"></param>
        static void ProcessPlaceholder(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null)
                throw new ArgumentNullException(nameof(openXmlElementDataContext));

            if (openXmlElementDataContext.Element.IsContentControl())
            {
                var element = openXmlElementDataContext.Element as SdtElement;
                var tagValue = GetTagValue(element);

                Match match = command.Match(tagValue);
                if (match.Success)
                {
                    Debug.WriteLine("FoundCommand: " + tagValue);

                    var commandName = match.Groups["command"].Value;
                    var commandParams = match.Groups["params"].Value;

                    var properties = new List<string>();
                    if (!string.IsNullOrEmpty(commandParams))
                    {
                        properties = commandParams
                            .Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(it => it.Trim(new char[] { '"', ' ' }))
                            .ToList();
                    }

                    var result = CommandHandler.Handle(commandName, properties, openXmlElementDataContext.DataContext);
                    if (result.OutputType == CommandOutputType.Content)
                    {
                        if (testPropertyValues.TryGetValue(tagValue, out var values))
                            values.Add(result.OutputContent);
                        else
                            testPropertyValues.Add(tagValue, new() { result.OutputContent });

                        SetContentOfContentControl(openXmlElementDataContext.Element as SdtElement, result.OutputContent);
                    }
                    else if (result.OutputType == CommandOutputType.None)
                    {
                        if (result.DataContext != null)
                            PopulateOtherOpenXmlElements(new(openXmlElementDataContext.Element, result.DataContext));
                        else
                            openXmlElementDataContext.Element.Remove();
                    }
                    else if (result.OutputType == CommandOutputType.List)
                    {
                        foreach (object item in result.OutputList)
                            CloneElementAndSetContentInPlaceholders(new(openXmlElementDataContext.Element, item));
                        openXmlElementDataContext.Element.Remove();
                    }
                }
            }
            else
                PopulateOtherOpenXmlElements(openXmlElementDataContext.CloneTyped());
        }

        static void SetContentOfContentControl(SdtElement element, string content)
        {
            // Set text without data binding
            OpenXmlHelper.SetContentOfContentControl(element, content);
        }

        /// <summary>
        /// Cloned element and set content in placeholders.
        /// </summary>
        /// <param name="openXmlElementDataContext"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="NullReferenceException"></exception>
        static void CloneElementAndSetContentInPlaceholders(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null)
                throw new ArgumentNullException(nameof(openXmlElementDataContext));

            if (openXmlElementDataContext.Element == null)
                throw new NullReferenceException(nameof(openXmlElementDataContext.Element));

            SdtElement clonedSdtElement;
            if (openXmlElementDataContext.Element.Parent != null && openXmlElementDataContext.Element.Parent is Paragraph)
            {
                Paragraph clonedPara = openXmlElementDataContext.Element.Parent.InsertBeforeSelf(openXmlElementDataContext.Element.Parent.CloneNode(true) as Paragraph);
                clonedSdtElement = clonedPara.Descendants<SdtElement>().First();
            }
            else
            {
                clonedSdtElement = openXmlElementDataContext.Element.InsertBeforeSelf(openXmlElementDataContext.Element.CloneNode(true) as SdtElement);
            }

            foreach (var v in clonedSdtElement.Elements())
                ProcessPlaceholder(new(v, openXmlElementDataContext.DataContext));
        }

        /// <summary>
        /// Populates the other open XML elements.
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        static void PopulateOtherOpenXmlElements(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext.Element is OpenXmlCompositeElement && openXmlElementDataContext.Element.HasChildren)
            {
                List<OpenXmlElement> elements = openXmlElementDataContext.Element.Elements().ToList();

                foreach (var element in elements)
                {
                    if (element is OpenXmlCompositeElement)
                        ProcessPlaceholder(new(element, openXmlElementDataContext.DataContext));
                }
            }
        }

        /// <summary>
        /// Gets the tag value.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        static string GetTagValue(SdtElement element)
        {
            Tag tag = OpenXmlHelper.GetTag(element);

            return (tag == null || (tag.Val.HasValue == false)) ? string.Empty : tag.Val.Value;
        }
    }
}