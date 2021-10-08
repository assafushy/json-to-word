using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordParagraph : IWordObject
    {
        public int HeadingLevel { get; set; }
        public List<WordRun> Runs { get; set; }
        public WordObjectType Type { get; set; }
    }
}