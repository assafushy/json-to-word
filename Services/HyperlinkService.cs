using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace JsonToWord.Services
{
    internal class HyperlinkService
    {
        internal static Hyperlink CreateHyperlink(string id)
        {
            return new Hyperlink() { History = true, Id = id };
        }

        internal static string AddHyperlinkRelationship(MainDocumentPart mainDocumentPart, Uri uri)
        {
            var id = CreateHyperlinkId();
            mainDocumentPart.AddHyperlinkRelationship(uri, true, id);

            return id;
        }
        private static string CreateHyperlinkId()
        {
            return "relId" + Guid.NewGuid().ToString("N");
        }
    }
}