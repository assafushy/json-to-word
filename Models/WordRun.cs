namespace JsonToWord.Models
{
    public class WordRun
    {
        public WordAttachment Attachment { get; set; }
        public string Font { get; set; }
        public bool InsertLineBreak { get; private set; } = false;
        public bool InsertSpace { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public int Size { get; set; }
        public string Text { get; set; }
        public string Uri { get; set; }
        public string FontColor { get; set; }

        public WordRun()
        {
            Font = "Arial";
            Size = 12;

        }
    }
}