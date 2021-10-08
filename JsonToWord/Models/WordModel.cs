using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordModel
    {
        public string TemplatePath { get; set; }
        public List<WordContentControl> ContentControls { get; set; }
    }
}