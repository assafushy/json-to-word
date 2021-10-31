using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Services.Interfaces;

namespace JsonToWord.Services
{
    internal class ContentControlService
    {
        internal void ClearContentControl(WordprocessingDocument document, string contentControlTitle, bool force)
        {
            var sdtBlock = document.MainDocumentPart.Document.Body.Descendants<SdtBlock>()
                .FirstOrDefault(e => e.Descendants<SdtAlias>().FirstOrDefault()?.Val == contentControlTitle);

            if (sdtBlock == null)
                throw new Exception("Did not find a content control with the title " + contentControlTitle);

            if (!string.IsNullOrEmpty(sdtBlock.InnerText) && sdtBlock.InnerText == "Click or tap here to enter text." || force)
                RemoveAllStdContentBlock(sdtBlock);
        }

        internal SdtBlock FindContentControl(WordprocessingDocument preprocessingDocument, string contentControlTitle)
        {
            var sdtBlock = preprocessingDocument.MainDocumentPart.Document.Body.Descendants<SdtBlock>().FirstOrDefault(e => e.Descendants<SdtAlias>().FirstOrDefault()?.Val == contentControlTitle);

            if (sdtBlock == null)
                throw new Exception("Did not find a content control with the title " + contentControlTitle);

            return sdtBlock;
        }

        internal void RemoveContentControl(WordprocessingDocument document, string contentControlTitle)
        {
            var contentControl = FindContentControl(document, contentControlTitle);

            foreach (var element in contentControl.Elements())
            {
                if (element is SdtContentBlock)
                {
                    foreach (var innerElement in element.Elements())
                    {
                        contentControl.Parent.InsertBefore(innerElement.CloneNode(true), contentControl);
                    }
                }
            }

            contentControl.Remove();
        }

        private void RemoveAllStdContentBlock(SdtBlock sdtBlock)
        {
            var childElements = new List<OpenXmlElement>();

            foreach (var childElement in sdtBlock.ChildElements)
            {
                if (childElement is SdtContentBlock)
                {
                    childElements.Add(childElement);
                }
            }

            foreach (var childElement in childElements)
            {
                sdtBlock.RemoveChild(childElement);
            }
        }
    }
}