namespace JsonToWord.Models
{
    internal class DrawingExtent
    {
        internal int Height { get; set; }
        internal int Width { get; set; }

        internal DrawingExtent(int height, int width)
        {
            Height = height;
            Width = width;
        }
    }
}