using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace BrandUp.DocumentTemplater.Internals
{
    /// <summary>
    /// Helper class for OpenXml operations for document generation
    /// </summary>
    internal static class OpenXmlHelper
    {
        readonly static Random rng = new();

        #region Public Methods

        /// <summary>
        /// Gets the SDT content of content control.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        public static OpenXmlCompositeElement GetSdtContentOfContentControl(SdtElement element)
        {
            if (element is SdtRun sdtRunELement)
            {
                return sdtRunELement.SdtContentRun;
            }
            else if (element is SdtBlock sdtBlockElement)
            {
                return sdtBlockElement.SdtContentBlock;
            }
            else if (element is SdtCell sdtCellElement)
            {
                return sdtCellElement.SdtContentCell;
            }
            else if (element is SdtRow sdtRowElement)
            {
                return sdtRowElement.SdtContentRow;
            }

            return null;
        }

        /// <summary>
        /// Ensures the unique content control ids for main document part.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        public static void EnsureUniqueContentControlIdsForMainDocumentPart(MainDocumentPart mainDocumentPart)
        {
            List<int> contentControlIds = new();

            if (mainDocumentPart != null)
            {
                foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                {
                    SetUniquecontentControlIds(part.Header, contentControlIds);
                    part.Header.Save();
                }

                foreach (FooterPart part in mainDocumentPart.FooterParts)
                {
                    SetUniquecontentControlIds(part.Footer, contentControlIds);
                    part.Footer.Save();
                }

                SetUniquecontentControlIds(mainDocumentPart.Document.Body, contentControlIds);
                mainDocumentPart.Document.Save();
            }
        }

        /// <summary>
        /// Sets the unique content control ids.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="existingIds">The existing ids.</param>
        public static void SetUniquecontentControlIds(OpenXmlCompositeElement element, List<int> existingIds)
        {
            foreach (SdtId sdtId in element.Descendants<SdtId>())
            {
                if (existingIds.Contains(sdtId.Val))
                {
                    int randomId = rng.Next(int.MaxValue);

                    while (existingIds.Contains(randomId))
                    {
                        rng.Next(int.MaxValue);
                    }

                    sdtId.Val.Value = randomId;
                }
                else
                {
                    existingIds.Add(sdtId.Val);
                }
            }
        }

        /// <summary>
        /// Sets the content of content control.
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <param name="content">The content.</param>
        public static void SetContentOfContentControl(SdtElement contentControl, string content)
        {
            if (contentControl == null)
                throw new ArgumentNullException(nameof(contentControl));

            content = string.IsNullOrEmpty(content) ? string.Empty : content;
            bool isCombobox = contentControl.SdtProperties.Descendants<SdtContentDropDownList>().FirstOrDefault() != null;

            if (isCombobox)
            {
                OpenXmlCompositeElement openXmlCompositeElement = GetSdtContentOfContentControl(contentControl);
                Run run = CreateRun(openXmlCompositeElement, content);
                SetSdtContentKeepingPermissionElements(openXmlCompositeElement, run);
            }
            else
            {
                OpenXmlCompositeElement openXmlCompositeElement = GetSdtContentOfContentControl(contentControl);
                contentControl.SdtProperties.RemoveAllChildren<ShowingPlaceholder>();
                List<Run> runs = new();

                if (IsContentControlMultiline(contentControl))
                {
                    List<string> textSplitted = content.Split(Environment.NewLine.ToCharArray()).ToList();
                    bool addBreak = false;

                    foreach (string textSplit in textSplitted)
                    {
                        Run run = CreateRun(openXmlCompositeElement, textSplit);

                        if (addBreak)
                        {
                            run.AppendChild<Break>(new Break());
                        }

                        if (!addBreak)
                        {
                            addBreak = true;
                        }

                        runs.Add(run);
                    }
                }
                else
                {
                    runs.Add(CreateRun(openXmlCompositeElement, content));
                }

                if (openXmlCompositeElement is SdtContentCell)
                {
                    AddRunsToSdtContentCell(openXmlCompositeElement as SdtContentCell, runs);
                }
                else if (openXmlCompositeElement is SdtContentBlock)
                {
                    Paragraph para = CreateParagraph(openXmlCompositeElement, runs);
                    SetSdtContentKeepingPermissionElements(openXmlCompositeElement, para);
                }
                else
                {
                    SetSdtContentKeepingPermissionElements(openXmlCompositeElement, runs);
                }
            }
        }

        /// <summary>
        /// Gets the tag.
        /// </summary>
        /// <param name="element">The SDT element.</param>
        /// <returns>
        /// Returns Tag of content control
        /// </returns>
        public static Tag GetTag(SdtElement element)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(element));

            return element.SdtProperties.Elements<Tag>().FirstOrDefault();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Determines whether [is content control multiline] [the specified content control].
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <returns>
        ///   <c>true</c> if [is content control multiline] [the specified content control]; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsContentControlMultiline(SdtElement contentControl)
        {
            SdtContentText contentText = contentControl.SdtProperties.Elements<SdtContentText>().FirstOrDefault();

            bool isMultiline = false;

            if (contentText != null && contentText.MultiLine != null)
            {
                isMultiline = contentText.MultiLine.Value == true;
            }
            return isMultiline;
        }

        /// <summary>
        /// Sets the SDT content keeping permission elements.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="newChild">The new child.</param>
        private static void SetSdtContentKeepingPermissionElements(OpenXmlCompositeElement openXmlCompositeElement, OpenXmlElement newChild)
        {
            PermStart start = openXmlCompositeElement.Descendants<PermStart>().FirstOrDefault();
            PermEnd end = openXmlCompositeElement.Descendants<PermEnd>().FirstOrDefault();
            openXmlCompositeElement.RemoveAllChildren();

            if (start != null)
            {
                openXmlCompositeElement.AppendChild(start);
            }

            openXmlCompositeElement.AppendChild(newChild);

            if (end != null)
            {
                openXmlCompositeElement.AppendChild(end);
            }
        }

        /// <summary>
        /// Sets the SDT content keeping permission elements.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="newChildren">The new children.</param>
        private static void SetSdtContentKeepingPermissionElements(OpenXmlCompositeElement openXmlCompositeElement, List<Run> newChildren)
        {
            PermStart start = openXmlCompositeElement.Descendants<PermStart>().FirstOrDefault();
            PermEnd end = openXmlCompositeElement.Descendants<PermEnd>().FirstOrDefault();
            openXmlCompositeElement.RemoveAllChildren();

            if (start != null)
            {
                openXmlCompositeElement.AppendChild(start);
            }

            foreach (var newChild in newChildren)
            {
                openXmlCompositeElement.AppendChild(newChild);
            }

            if (end != null)
            {
                openXmlCompositeElement.AppendChild(end);
            }
        }

        /// <summary>
        /// Adds the runs to SDT content cell.
        /// </summary>
        /// <param name="sdtContentCell">The SDT content cell.</param>
        /// <param name="runs">The runs.</param>
        private static void AddRunsToSdtContentCell(SdtContentCell sdtContentCell, List<Run> runs)
        {
            TableCell cell = new();
            Paragraph para = new();
            para.RemoveAllChildren();

            foreach (Run run in runs)
            {
                para.AppendChild(run);
            }

            cell.AppendChild<Paragraph>(para);
            SetSdtContentKeepingPermissionElements(sdtContentCell, cell);
        }

        /// <summary>
        /// Creates the paragraph.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="runs">The runs.</param>
        /// <returns></returns>
        private static Paragraph CreateParagraph(OpenXmlCompositeElement openXmlCompositeElement, List<Run> runs)
        {
            ParagraphProperties paragraphProperties = openXmlCompositeElement.Descendants<ParagraphProperties>().FirstOrDefault();
            Paragraph para;
            if (paragraphProperties != null)
            {
                para = new Paragraph(paragraphProperties.CloneNode(true));
                foreach (Run run in runs)
                {
                    para.AppendChild<Run>(run);
                }
            }
            else
            {
                para = new Paragraph();
                foreach (Run run in runs)
                {
                    para.AppendChild<Run>(run);
                }
            }
            return para;
        }

        /// <summary>
        /// Creates the run.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="content">The content.</param>
        /// <returns></returns>
        private static Run CreateRun(OpenXmlCompositeElement openXmlCompositeElement, string content)
        {
            RunProperties runProperties = openXmlCompositeElement.Descendants<RunProperties>().FirstOrDefault();

            Run run;
            if (runProperties != null)
            {
                run = new Run(runProperties.CloneNode(true), new Text(content));
            }
            else
            {
                run = new Run(new Text(content));
            }

            return run;
        }

        #endregion
    }
}
