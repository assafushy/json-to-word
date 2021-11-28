using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordTable : IWordObject
    {
        public WordObjectType Type { get; set; }
        public List<WordTableRow> Rows { get; set; }
        public bool RepeatHeaderRow { get; set; }
    }
}