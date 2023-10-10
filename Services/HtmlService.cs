using System;
using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using JsonToWord.Models;

namespace JsonToWord.Services
{
    internal class HtmlService
    {
        private readonly ContentControlService _contentControlService;
        public HtmlService()
        {
            _contentControlService = new ContentControlService();
        }
        internal void Insert(WordprocessingDocument document, string contentControlTitle, WordHtml wordHtml)
        {
            var html = SetHtmlFormat(wordHtml.Html);
            
            html = RemoveWordHeading(html);

            html = FixBullets(html);

            var tempHtmlFile = CreateHtmlWordDocument(html);

            var altChunkId = "altChunkId" + Guid.NewGuid().ToString("N");
            var mainPart = document.MainDocumentPart;
            var chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            using (var fileStream = File.Open(tempHtmlFile, FileMode.Open))
            {
                chunk.FeedData(fileStream);
            }

            var altChunk = new AltChunk { Id = altChunkId };
            
            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(altChunk);

            sdtBlock.AppendChild(sdtContentBlock);
        }

        internal string CreateHtmlWordDocument(string html)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            var tempHtmlDirectory = Path.Combine(Path.GetTempPath(), "MicrosoftWordOpenXml", Guid.NewGuid().ToString("N"));

            if (!Directory.Exists(tempHtmlDirectory))
                Directory.CreateDirectory(tempHtmlDirectory);

            var tempHtmlFile = CreateTempDocument(tempHtmlDirectory);

            using (var document = WordprocessingDocument.Open(tempHtmlFile, true))
            {
                var mainPart = document.MainDocumentPart;

                if (mainPart == null)
                {
                    mainPart = document.AddMainDocumentPart();
                    new Document(new Body()).Save(mainPart);
                }

                var converter = new HtmlConverter(mainPart);
                try
                {
                    converter.ParseHtml(html);
                }
                catch(Exception ex)
                {
                    log.Error("DocGen ran into an issue parsing the html due to :" , ex);
                    converter.ParseHtml("<p style='color: red'><b>DocGen ran into an issue parsing the html due to :" + ex.Message +"<b></p>");
                }
            }

            return tempHtmlFile;
        }

        private string CreateTempDocument(string directory)
        {
            var tempDocumentFile = Path.Combine(directory, $"{Guid.NewGuid():N}.docx");

            using (var wordDocument = WordprocessingDocument.Create(tempDocumentFile, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDocument.AddMainDocumentPart();

                mainPart.Document = new Document();
                mainPart.Document.AppendChild(new Body());
            }

            return tempDocumentFile;
        }

        private string SetHtmlFormat(string html)
        {
            if (!html.ToLower().StartsWith("<html>"))
                return $"<html style=\"font-family: Arial, sans-serif; font-size: 12pt;\"><body>{html}</body></html>";

            return html;
        }

        private string RemoveWordHeading(string html)
        {
            var result = Regex.Replace(html, @"(?s)<h\d.+?>", string.Empty);
            return Regex.Replace(result, @"</h\d>", string.Empty);
        }

        private string FixBullets(string html)
        {
            html = FixBullets(html, "MsoListParagraphCxSpFirst");
            html = FixBullets(html, "MsoListParagraphCxSpMiddle");
            html = FixBullets(html, "MsoListParagraphCxSpLast");

            return html;
        }

        private static string FixBullets(string description, string mainClassPattern)
        {
            var res = description;

            foreach (var match in Regex.Matches(description, $"(?s)<p class={mainClassPattern}.*?</p>", RegexOptions.IgnoreCase))
            {
                var bulletPattern = "(?s)<span style=\"font-family:Symbol;\">.*?</span></span></span>";

                var bulletMatch = Regex.Match(match.ToString(), bulletPattern, RegexOptions.IgnoreCase);

                if (!bulletMatch.Success)
                    continue;

                var matchWithoutBullet = Regex.Replace(match.ToString(), bulletPattern, string.Empty);

                var innerMatch = Regex.Match(matchWithoutBullet, "(?=>)(.*?)(?=</p>)", RegexOptions.Singleline);

                if (!innerMatch.Success)
                    continue;

                var newText = matchWithoutBullet.Replace(innerMatch.Value, $"><ul><li>{innerMatch.Value.Remove(0, 1)}</li></ul>");

                res = res.Replace(match.ToString(), newText);
            }

            return res;
        }
    }
}