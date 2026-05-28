using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using BrandUp.DocumentTemplater.Handling;
using BrandUp.DocumentTemplater.Internals;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
        public static async Task<Stream> GenerateDocumentAsync(object dataContext, Stream templateStream, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(dataContext);
            ArgumentNullException.ThrowIfNull(templateStream);

            // WordprocessingDocument изменяет поток, поэтому сначала создаем выходной поток. 
            var output = new MemoryStream();
            await templateStream.CopyToAsync(output, cancellationToken);

            using (var wordDocument = WordprocessingDocument.Open(output, true))
            {
                wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
                Document document = mainDocumentPart.Document;

                foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                    ProcessPlaceholder(new(part.Header, dataContext));

                cancellationToken.ThrowIfCancellationRequested();

                foreach (FooterPart part in mainDocumentPart.FooterParts)
                    ProcessPlaceholder(new(part.Footer, dataContext));

                cancellationToken.ThrowIfCancellationRequested();

                ProcessPlaceholder(new(document, dataContext));

                // Сохранение всех частей выполняется здесь: метод проставляет
                // уникальные id и вызывает Save() для header/footer/документа.
                OpenXmlHelper.EnsureUniqueContentControlIdsForMainDocumentPart(mainDocumentPart);
            }

            output.Seek(0, SeekOrigin.Begin);
            return output;
        }

        #region Helpers

        /// <summary>
        /// Разбирает строку параметров команды на отдельные значения.
        /// Запятые внутри кавычек (' или ") не считаются разделителями,
        /// что позволяет использовать форматы вроде "#,##0.00".
        /// </summary>
        /// <param name="raw">Содержимое скобок команды.</param>
        /// <returns>Список параметров без обрамляющих кавычек и пробелов.</returns>
        internal static List<string> ParseParameters(string raw)
        {
            var result = new List<string>();
            if (string.IsNullOrEmpty(raw))
                return result;

            var current = new StringBuilder();
            char quote = '\0';
            bool inQuotes = false;

            foreach (char c in raw)
            {
                if (inQuotes)
                {
                    if (c == quote)
                        inQuotes = false;
                    else
                        current.Append(c);
                }
                else if (c == '"' || c == '\'')
                {
                    inQuotes = true;
                    quote = c;
                }
                else if (c == ',')
                {
                    AddParameter(result, current);
                    current.Clear();
                }
                else
                    current.Append(c);
            }

            AddParameter(result, current);
            return result;
        }

        static void AddParameter(List<string> result, StringBuilder builder)
        {
            var value = builder.ToString().Trim();
            if (value.Length > 0)
                result.Add(value);
        }

        /// <summary>
        /// Обрабатывает заглушку
        /// </summary>
        /// <param name="openXmlElementDataContext">Контекст данных элемента "open XML".</param>
        static void ProcessPlaceholder(OpenXmlElementDataContext openXmlElementDataContext)
        {
            ArgumentNullException.ThrowIfNull(openXmlElementDataContext);

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

                    var properties = ParseParameters(commandParams);

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
                PopulateOtherOpenXmlElements(openXmlElementDataContext);
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
            ArgumentNullException.ThrowIfNull(openXmlElementDataContext);

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
                List<OpenXmlElement> elements = [.. openXmlElementDataContext.Element.Elements()];

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

            return (tag?.Val == null || !tag.Val.HasValue) ? string.Empty : tag.Val.Value;
        }

        #endregion
    }
}