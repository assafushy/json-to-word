using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using System;
using System.Collections.Generic;  // Ensure this using directive is included for List<T>

namespace JsonToWord.Services
{
    internal class ParagraphService
    {
        internal Paragraph CreateParagraph(WordParagraph wordParagraph)
        {
            var paragraph = new Paragraph();
  
            if (wordParagraph.HeadingLevel == 0)
                return paragraph;
foreach (var wordRun in wordParagraph.Runs)
{
    string runDetails = $"Text: {wordRun.Text}, " +
                        $"Bold: {wordRun.Bold}, " +
                        $"Italic: {wordRun.Italic}, " +
                        $"Underline: {wordRun.Underline}, " +
                        $"Size: {wordRun.Size}, " +
                        $"Font: {wordRun.Font}, " +
                        $"Uri: {wordRun.Uri}, " +
                        $"FontColor: {wordRun.FontColor}, " +
                        $"Attachment: {wordRun.Attachment}";

    Console.WriteLine("-------------Run Details----------: " + runDetails);
}

            var paragraphProperties = new ParagraphProperties();
            var paragraphStyleId = new ParagraphStyleId { Val = $"Heading{wordParagraph.HeadingLevel}" };

            paragraphProperties.AppendChild(paragraphStyleId);
            paragraph.AppendChild(paragraphProperties);

            return paragraph;
        }
    }
}