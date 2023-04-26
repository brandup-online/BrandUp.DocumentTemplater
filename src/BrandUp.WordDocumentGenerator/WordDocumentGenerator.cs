using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;

namespace BrandUp.DocxGenerator
{
    public static class WordDocumentGenerator
    {
        readonly static Regex commandRegex = new(@"\{(?<command>\w+)\((?<params>.*)\)\}");
        readonly static OpenXmlHelper openXmlHelper = new("http://schemas.WordDocumentGenerator.com/DocumentGeneration");

        public static Stream GenerateDocument(object dataContext, byte[] templateBytes)
        {
            using var templateStream = new MemoryStream(templateBytes);
            return GenerateDocument(dataContext, templateStream);
        }

        public static Stream GenerateDocument(object dataContext, Stream templateStream)
        {
            if (dataContext == null)
                throw new ArgumentNullException(nameof(dataContext));
            if (templateStream == null)
                throw new ArgumentNullException(nameof(templateStream));

            using (var wordDocument = WordprocessingDocument.Open(templateStream, true))
            {
                wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);
                MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
                Document document = mainDocumentPart.Document;

                foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                {
                    ProcessPlaceholder(new OpenXmlElementDataContext(part.Header, dataContext));
                    part.Header.Save();
                }

                foreach (FooterPart part in mainDocumentPart.FooterParts)
                {
                    ProcessPlaceholder(new OpenXmlElementDataContext(part.Footer, dataContext));
                    part.Footer.Save();
                }

                ProcessPlaceholder(new OpenXmlElementDataContext(document, dataContext));
                openXmlHelper.EnsureUniqueContentControlIdsForMainDocumentPart(mainDocumentPart);
                document.Save();
            }

            var output = new MemoryStream();
            templateStream.Seek(0, SeekOrigin.Begin);
            templateStream.CopyTo(output);
            output.Seek(0, SeekOrigin.Begin);

            return output;
        }

        static void ProcessPlaceholder(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (IsContentControl(openXmlElementDataContext))
            {
                var element = openXmlElementDataContext.Element as SdtElement;
                var tagValue = GetTagValue(element);

                Match m = commandRegex.Match(tagValue);
                if (m.Success)
                {
                    Debug.WriteLine("FoundCommand: " + tagValue);

                    var commandName = m.Groups["command"].Value;
                    var commandParams = m.Groups["params"].Value;

                    var properties = new List<string>();
                    if (!string.IsNullOrEmpty(commandParams))
                    {
                        properties = commandParams.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        properties = properties.Select(it => it.Trim(new char[] { '"', ' ' })).ToList();
                    }

                    var e = FoundCommandEventArgs(commandName, properties, openXmlElementDataContext.DataContext);
                    if (e.OutputType == DocumentGeneratorCommandOutputType.Content)
                    {
                        SetContentOfContentControl(openXmlElementDataContext.Element as SdtElement, e.OutputContent);
                    }
                    else if (e.OutputType == DocumentGeneratorCommandOutputType.None)
                    {
                        if (e.DataContext != null)
                            PopulateOtherOpenXmlElements(new OpenXmlElementDataContext(openXmlElementDataContext.Element, e.DataContext));
                        else
                            openXmlElementDataContext.Element.Remove();
                    }
                    else if (e.OutputType == DocumentGeneratorCommandOutputType.List)
                    {
                        foreach (object item in e.OutputList)
                            CloneElementAndSetContentInPlaceholders(new OpenXmlElementDataContext(openXmlElementDataContext.Element, item));
                        openXmlElementDataContext.Element.Remove();
                    }
                }
            }
            else
                PopulateOtherOpenXmlElements(openXmlElementDataContext.CloneTyped());
        }

        static FoundCommandEventArgs FoundCommandEventArgs(string commandName, List<string> properties, object dataContext)
        {
            if (commandName == null)
                throw new ArgumentNullException(nameof(commandName));
            if (properties == null)
                throw new ArgumentNullException(nameof(properties));
            if (dataContext == null)
                throw new ArgumentNullException(nameof(dataContext));

            switch (commandName.ToLower())
            {
                case "setcontextofproperty":
                    {
                        Type t = dataContext.GetType();

                        var propName = properties[0];

                        object value = null;

                        bool isDict = t.GetInterfaces().Any(it => it.IsGenericType && it.GetGenericTypeDefinition() == typeof(IDictionary<,>));
                        if (isDict)
                            value = ((IDictionary<string, object>)dataContext)[propName];
                        else
                        {
                            PropertyInfo p = null;
                            try
                            {
                                p = t.GetProperty(propName);
                                //var props = t.GetProperties(System.Reflection.BindingFlags.FlattenHierarchy | System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | BindingFlags.Public);
                                //var p = props.First(x => x.Name == propName);
                                if (p == null)
                                    throw new InvalidOperationException("Не найдено свойство " + properties[0] + " у обхекта с типом " + t.FullName + ".");

                                value = p.GetValue(dataContext);
                            }
                            catch (Exception)
                            {
                                value = null;
                            }
                        }

                        return new(commandName, properties, value)
                        {
                            OutputType = DocumentGeneratorCommandOutputType.None
                        };
                    }
                case "foreach":
                    {
                        var items = new List<object>();
                        System.Collections.IEnumerable collection = dataContext as System.Collections.IEnumerable;
                        if (collection != null)
                        {
                            foreach (object item in collection)
                                items.Add(item);
                        }

                        return new(commandName, properties, dataContext)
                        {
                            OutputType = DocumentGeneratorCommandOutputType.List,
                            OutputList = items
                        };
                    }
                case "datetimenow":
                    {
                        string output;
                        DateTime d = DateTime.Now;

                        if (properties.Count > 0)
                            output = d.ToString(properties[0]);
                        else
                            output = d.ToString();

                        return new(commandName, properties, dataContext)
                        {
                            OutputType = DocumentGeneratorCommandOutputType.Content,
                            OutputContent = output
                        };
                    }
                case "prop":
                    {
                        Type t = dataContext.GetType();

                        var propName = properties[0];
                        string output = null;
                        object value = null;

                        bool isDict = t.GetInterfaces().Any(it => it.IsGenericType && it.GetGenericTypeDefinition() == typeof(IDictionary<,>));
                        if (isDict)
                        {
                            value = ((IDictionary<string, object>)dataContext)[propName];
                        }
                        else
                        {
                            try
                            {
                                var p = t.GetProperty(propName);
                                if (p == null)
                                    throw new InvalidOperationException("Не найдено свойство " + properties[0] + " у обхекта с типом " + t.FullName + ".");

                                value = p.GetValue(dataContext);

                                if (value == null)
                                    value = "";
                            }
                            catch (Exception)
                            {
                                value = "";
                            }
                        }
                        dataContext = value;

                        if (value != null)
                        {
                            if (properties.Count > 1 && !string.IsNullOrEmpty(properties[1]))
                            {
                                var format = properties[1];

                                if (value is DateTime)
                                    output = ((DateTime)value).ToString(format);
                                else if (value is TimeSpan)
                                    output = ((TimeSpan)value).ToString(format);
                                else if (value is decimal)
                                    output = ((decimal)value).ToString(format);
                                else if (value is double)
                                    output = ((double)value).ToString(format);
                                else if (value is float)
                                    output = ((float)value).ToString(format);
                                else if (value is int)
                                    output = ((int)value).ToString(format);
                                else if (value is bool)
                                {
                                    if (format == "b")
                                        output = (bool)value ? "да" : "нет";
                                    else if (format == "B")
                                        output = (bool)value ? "Да" : "Нет";
                                    else
                                        output = value.ToString();
                                }
                                else
                                    output = string.Format(format, value);
                            }
                            else
                                output = value.ToString();
                        }

                        return new(commandName, properties, dataContext)
                        {
                            OutputType = DocumentGeneratorCommandOutputType.Content,
                            OutputContent = output
                        };
                    }
                default: throw new NotSupportedException();
            }
        }
        static void SetContentOfContentControl(SdtElement element, string content)
        {
            // Set text without data binding
            openXmlHelper.SetContentOfContentControl(element, content);
        }

        static void CloneElementAndSetContentInPlaceholders(OpenXmlElementDataContext openXmlElementDataContext)
        {
            if (openXmlElementDataContext == null)
                throw new ArgumentNullException("openXmlElementDataContext");

            if (openXmlElementDataContext.Element == null)
                throw new ArgumentNullException("openXmlElementDataContext.element");

            SdtElement clonedSdtElement = null;

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
                ProcessPlaceholder(new OpenXmlElementDataContext(v, openXmlElementDataContext.DataContext));
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
                        ProcessPlaceholder(new OpenXmlElementDataContext(element, openXmlElementDataContext.DataContext));
                }
            }
        }

        /// <summary>
        /// Gets the tag value.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="templateTagPart">The template tag part.</param>
        /// <param name="tagGuidPart">The tag GUID part.</param>
        /// <returns></returns>
        static string GetTagValue(SdtElement element)
        {
            Tag tag = openXmlHelper.GetTag(element);

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