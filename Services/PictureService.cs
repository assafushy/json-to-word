using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using SixLabors.ImageSharp;
using System;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace JsonToWord.Services
{
    internal class PictureService
    {
        private readonly ContentControlService _contentControlService;
        public PictureService()
        {
            _contentControlService = new ContentControlService();
        }
        internal void Insert(WordprocessingDocument document, string contentControlTitle, WordAttachment wordAttachment)
        {
            var drawing = CreateDrawing(document.MainDocumentPart, wordAttachment.Path);

            var run = new Run();
            run.AppendChild(drawing);

            var paragraph = new Paragraph();
            paragraph.AppendChild(run);

            var sdtBlock = _contentControlService.FindContentControl(document, contentControlTitle);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(paragraph);

            sdtBlock.AppendChild(sdtContentBlock);
        }

        internal Drawing CreateDrawing(MainDocumentPart mainDocumentPart, string filePath) //int width = 3291840, int height = 2633473
        {
            var imagePartId = AddImagePart(mainDocumentPart, filePath);

            var drawingExtend = GetDrawingExtend(filePath);

            var drawing =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = drawingExtend.Width, Cy = drawingExtend.Height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(new A.GraphicData(new PIC.Picture(new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = imagePartId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = drawingExtend.Width, Cy = drawingExtend.Height }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                         { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            return drawing;
        }

        private static DrawingExtent GetDrawingExtend(string localPath)
        {
            int width;
            int height;
            using (var bmp = Image.Load(localPath))
            {
                width = bmp.Width;
                height = bmp.Height;
            }

            width = (int)Math.Round((decimal)width * 9525);
            height = (int)Math.Round((decimal)height * 9525);

            if (width > 5715000)
                width = 5715000;

            return new DrawingExtent(height, width);
        }

        private string AddImagePart(MainDocumentPart mainDocumentPart, string imagePath)
        {
            var imagePart = mainDocumentPart.AddImagePart(ImagePartType.Jpeg);

            using (var stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            return mainDocumentPart.GetIdOfPart(imagePart);
        }
    }
}