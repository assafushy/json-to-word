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
            var runsToRemove = new List<WordRun>(); // Declare the list here

            foreach (var wordRun in wordParagraph.Runs)
            {
                if (wordRun.Text == "Test Description:")
                {
                    runsToRemove.Add(wordRun);
                }
            }

            foreach (var runToRemove in runsToRemove)
            {
                wordParagraph.Runs.Remove(runToRemove);
            }
            
            if (wordParagraph.HeadingLevel == 0)
                return paragraph;
            foreach (var wordRun in wordParagraph.Runs)
            {
                Console.WriteLine("-------------Run Text----------: " + wordRun);
            }
            var paragraphProperties = new ParagraphProperties();
            var paragraphStyleId = new ParagraphStyleId { Val = $"Heading{wordParagraph.HeadingLevel}" };

            paragraphProperties.AppendChild(paragraphStyleId);
            paragraph.AppendChild(paragraphProperties);

            return paragraph;
        }
    }
}