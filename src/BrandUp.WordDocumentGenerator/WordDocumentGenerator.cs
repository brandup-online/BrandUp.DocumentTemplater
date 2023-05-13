using BrandUp.DocumentTemplater.Exeptions;
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

        /// <summary>
        /// Преобразует шаблон в .docx документ, записывая во все элементы управления соответствующие значения
        /// </summary>
        /// <param name="dataContext">Контекст данных</param>
        /// <param name="templateStream">Шаблон</param>
        /// <param name="cancellationToken">Токен отмены</param>
        /// <returns>.docx файл</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<Stream> GenerateDocument(object dataContext, Stream templateStream, CancellationToken cancellationToken)
        {
            if (dataContext == null)
                throw new DataContextNullException();
            if (templateStream == null)
                throw new ArgumentNullException(nameof(templateStream));

            // WordprocessingDocument изменяет поток, поэтому сначала создаем выходной поток. 
            var output = new MemoryStream();
            await templateStream.CopyToAsync(output, cancellationToken);

            using (var wordDocument = WordprocessingDocument.Open(output, true))
            {
                wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
                Document document = mainDocumentPart.Document;

                foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                {
                    ProcessPlaceholder(new(part.Header, dataContext));
                    part.Header.Save();
                }

                cancellationToken.ThrowIfCancellationRequested();

                foreach (FooterPart part in mainDocumentPart.FooterParts)
                {
                    ProcessPlaceholder(new(part.Footer, dataContext));
                    part.Footer.Save();
                }

                cancellationToken.ThrowIfCancellationRequested();

                ProcessPlaceholder(new(document, dataContext));
                OpenXmlHelper.EnsureUniqueContentControlIdsForMainDocumentPart(mainDocumentPart);
                document.Save();
            }

            output.Seek(0, SeekOrigin.Begin);
            return output;
        }

        #region Helpers

        /// <summary>
        /// Обрабатывает заглушку
        /// </summary>
        /// <param name="openXmlElementDataContext">Контекст данных элемента "open XML".</param>
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
        /// Клонирует элемент и записывает данные 
        /// </summary>
        /// <param name="openXmlElementDataContext">Контекст данных элемента "open XML".</param>
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
        /// Заполняет другие открытые элементы XML.
        /// </summary>
        /// <param name="openXmlElementDataContext">Контекст данных элемента "open XML".</param>
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
        /// Получает значение тега.
        /// </summary>
        /// <param name="element">Элемент.</param>
        /// <returns>тег</returns>
        static string GetTagValue(SdtElement element)
        {
            Tag tag = OpenXmlHelper.GetTag(element);

            return (tag == null || (tag.Val.HasValue == false)) ? string.Empty : tag.Val.Value;
        }

        #endregion
    }
}