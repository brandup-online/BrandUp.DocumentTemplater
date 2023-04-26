using BrandUp.DocxGenerator.Handling;
using BrandUp.DocxGenerator.Internals;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace BrandUp.DocxGenerator
{
    public static class WordDocumentGenerator
    {
        readonly static Regex command = new(@"\{(?<command>\w+)\((?<params>.*)\)\}", RegexOptions.IgnoreCase);
        public static async Task<Stream> GenerateDocument(object dataContext, byte[] templateBytes, CancellationToken cancellationToken)
        {
            using var templateStream = new MemoryStream(templateBytes);
            return await GenerateDocument(dataContext, templateStream, cancellationToken);
        }

        public static async Task<Stream> GenerateDocument(object dataContext, Stream templateStream, CancellationToken cancellationToken)
        {
            if (dataContext == null)
                throw new ArgumentNullException(nameof(dataContext));
            if (templateStream == null)
                throw new ArgumentNullException(nameof(templateStream));

            using (var wordDocument = WordprocessingDocument.Open(templateStream, true))
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

            var output = new MemoryStream();

            templateStream.Seek(0, SeekOrigin.Begin);
            await templateStream.CopyToAsync(output, cancellationToken);

            output.Seek(0, SeekOrigin.Begin);

            return output;
        }

        static void ProcessPlaceholder(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (IsContentControl(openXmlElementDataContext))
            {
                var element = openXmlElementDataContext.Element as SdtElement;
                var tagValue = GetTagValue(element);

                Match m = command.Match(tagValue);
                if (m.Success)
                {
                    Debug.WriteLine("FoundCommand: " + tagValue);

                    var commandName = m.Groups["command"].Value;
                    var commandParams = m.Groups["params"].Value;

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

            string fullTag = (tag == null || (tag.Val.HasValue == false)) ? string.Empty : tag.Val.Value;

            return fullTag;
        }

        /// <summary>
        /// Determines whether [is content control] [the specified open XML element data context].
        /// </summary>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        /// <returns>
        ///   <c>true</c> if [is content control] [the specified open XML element data context]; otherwise, <c>false</c>.
        /// </returns>
        static bool IsContentControl(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null || openXmlElementDataContext.Element == null)
                return false;

            return openXmlElementDataContext.Element is SdtBlock || openXmlElementDataContext.Element is SdtRun || openXmlElementDataContext.Element is SdtRow || openXmlElementDataContext.Element is SdtCell;
        }
    }
}