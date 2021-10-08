using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordContentControl
    {
        public bool ForceClean { get; set; }
        public string Title { get; set; }
        public List<IWordObject> WordObjects { get; set; }
    }
}