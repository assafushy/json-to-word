using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JsonToWord.Models;
using OVML = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;

namespace JsonToWord.Services
{
    internal class FileService
    {
        internal void Insert(WordprocessingDocument document, string contentControlTitle, WordAttachment wordAttachment)
        {
            var embeddedFile = CreateEmbeddedObject(document.MainDocumentPart, wordAttachment.Path, true);

            var run = new Run();
            run.AppendChild(embeddedFile);

            var paragraph = new Paragraph();
            paragraph.AppendChild(run);

            var sdtContentBlock = new SdtContentBlock();
            sdtContentBlock.AppendChild(paragraph);

            var contentControlService = new ContentControlService();
            var sdtBlock = contentControlService.FindContentControl(document, contentControlTitle);
            sdtBlock.AppendChild(sdtContentBlock);
        }

        internal EmbeddedObject CreateEmbeddedObject(MainDocumentPart mainDocumentPart, string filePath, bool displayAsIcon)
        {
            EmbeddedObject embeddedObject;

            var fileInfo = new FileInfo(filePath);
            var openXmlEmbeddedObject = new OpenXmlEmbeddedObject(fileInfo, displayAsIcon);


            if (string.IsNullOrEmpty(openXmlEmbeddedObject.OleObjectBinaryData))
                return null;

            using (var dataStream = new MemoryStream(Convert.FromBase64String(openXmlEmbeddedObject.OleObjectBinaryData)))
            {
                if (string.IsNullOrEmpty(openXmlEmbeddedObject.OleImageBinaryData))
                    return null;

                using (var emfStream = new MemoryStream(Convert.FromBase64String(openXmlEmbeddedObject.OleImageBinaryData)))
                {
                    var imagePartId = GetUniqueXmlItemId();
                    var imagePart = mainDocumentPart.AddImagePart(ImagePartType.Emf, imagePartId);

                    imagePart.FeedData(emfStream);

                    var embeddedPackagePartId = GetUniqueXmlItemId();

                    if (openXmlEmbeddedObject.ObjectIsOfficeDocument)
                    {
                        var embeddedObjectPart = mainDocumentPart.AddNewPart<EmbeddedPackagePart>(openXmlEmbeddedObject.FileContentType, embeddedPackagePartId);
                        embeddedObjectPart.FeedData(dataStream);
                    }
                    else
                    {
                        var embeddedObjectPart = mainDocumentPart.AddNewPart<EmbeddedObjectPart>(openXmlEmbeddedObject.FileContentType, embeddedPackagePartId);
                        embeddedObjectPart.FeedData(dataStream);
                    }

                    embeddedObject = GetEmbeddedObject(openXmlEmbeddedObject.FileType, imagePartId, openXmlEmbeddedObject.OleImageStyle, embeddedPackagePartId);
                }

            }

            return embeddedObject;
        }

        private static string GetUniqueXmlItemId()
        {
            return "r" + Guid.NewGuid().ToString().Replace("-", "");
        }

        private static EmbeddedObject GetEmbeddedObject(string fileType, string imageId, string imageStyle, string embeddedPackageId)
        {
            var embeddedObject = new EmbeddedObject();

            var shapeId = GetUniqueXmlItemId();
            var shape = new V.Shape() { Id = shapeId, Style = imageStyle };
            var imageData = new V.ImageData() { Title = "", RelationshipId = imageId };

            shape.AppendChild(imageData);
            var oleObject = new OVML.OleObject()
            {
                Type = OVML.OleValues.Embed,
                ProgId = fileType,
                ShapeId = shapeId,
                DrawAspect = OVML.OleDrawAspectValues.Icon,
                ObjectId = GetUniqueXmlItemId(),
                Id = embeddedPackageId
            };

            embeddedObject.AppendChild(shape);
            embeddedObject.AppendChild(oleObject);

            return embeddedObject;
        }
    }
}