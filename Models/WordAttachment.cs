namespace JsonToWord.Models
{
    public class WordAttachment : IWordObject
    {
        public string Path { get; set; }
        public WordObjectType Type { get; set; }
    }
}