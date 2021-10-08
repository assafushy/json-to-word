using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using System;

namespace JsonToWord.Services
{
    internal class TextService
    {
        internal void Write(WordprocessingDocument document, string contentControlTitle, WordParagraph wordParagraph)
        {
            var paragraphService = new ParagraphService();
            var paragraph = paragraphService.CreateParagraph(wordParagraph);

            var runService = new RunService();

            if (wordParagraph.Runs != null)
            {
                foreach (var wordRun in wordParagraph.Runs)
                {
                    var run = runService.CreateRun(wordRun);

                    if (wordRun.Uri != null)
                    {
                        var id = HyperlinkService.AddHyperlinkRelationship(document.MainDocumentPart, new Uri(wordRun.Uri));
                        var hyperlink = HyperlinkService.CreateHyperlink(id);
                        hyperlink.AppendChild(run);

                        paragraph.AppendChild(hyperlink);
                    }
                    else
                    {
                        paragraph.AppendChild(run);
                    }
                }
            }

            var contentControlService = new ContentControlService();
            var sdtBlock = contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(paragraph);

            sdtBlock.AppendChild(sdtContentBlock);
        }
    }
}