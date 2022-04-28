using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using System;

namespace JsonToWord.Services
{
    internal class TextService
    {
        private readonly ParagraphService _paragraphService;
        private readonly RunService _runService;
        private readonly ContentControlService _contentControlService;

        public TextService()
        {
            _paragraphService = new ParagraphService();
            _runService = new RunService();
            _contentControlService = new ContentControlService();
        }
        internal void Write(WordprocessingDocument document, string contentControlTitle, WordParagraph wordParagraph)
        {
            var paragraph = _paragraphService.CreateParagraph(wordParagraph);


            if (wordParagraph.Runs != null)
            {
                foreach (var wordRun in wordParagraph.Runs)
                {
                    var run = _runService.CreateRun(wordRun);

                    if (wordRun.Uri != null)
                    {
                        try
                        {
                        var id = HyperlinkService.AddHyperlinkRelationship(document.MainDocumentPart, new Uri(wordRun.Uri));
                        var hyperlink = HyperlinkService.CreateHyperlink(id);
                        hyperlink.AppendChild(run);

                        paragraph.AppendChild(hyperlink);
                        }
                        catch (UriFormatException e)
                        {
                            Console.WriteLine(wordRun.Uri+ " is an invalid uri \n" + e.Message);
                            paragraph.AppendChild(run);
                        }
                    }
                    else
                    {
                        paragraph.AppendChild(run);
                    }
                }
            }

            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(paragraph);

            sdtBlock.AppendChild(sdtContentBlock);
        }
    }
}