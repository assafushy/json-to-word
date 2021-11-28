using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordTableCell
    {
        //public bool GatherAttachment { get; set; }
        public List<WordParagraph> Paragraphs { get; set; }
        public WordShading Shading { get; set; }
        public string Width { get; set; }
        public List<WordAttachment> Attachments { get; set; }
        public WordHtml Html { get; set; }
    }
}