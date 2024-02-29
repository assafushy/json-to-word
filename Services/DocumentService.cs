using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace JsonToWord.Services
{
    public class DocumentService
    {
        public string CreateDocument(string templatePath)
        {
            var destinationFile = templatePath.Replace(".dot", ".doc");
            byte[] templateBytes = File.ReadAllBytes(templatePath);

            using (var templateStream = new MemoryStream())
            {
                templateStream.Write(templateBytes, 0, templateBytes.Length);

                using (var document = WordprocessingDocument.Open(templateStream, true))
                {
                    document.ChangeDocumentType(WordprocessingDocumentType.Document);
                    var mainPart = document.MainDocumentPart;

                    RemoveEmptyParagraphs(mainPart);
                    SetLandscape(mainPart);

                    mainPart.Document.Save();
                }

                File.WriteAllBytes(destinationFile, templateStream.ToArray());
            }

            return destinationFile;
        }

        private void RemoveEmptyParagraphs(MainDocumentPart mainPart)
        {
            var paragraphs = mainPart.Document.Body.Elements<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                if (!paragraph.Elements<Run>().Any(r => r.Elements<Text>().Any(t => !string.IsNullOrEmpty(t.Text.Trim()))))
                {
                    paragraph.Remove();
                }
            }
        }

        public void SetLandscape(MainDocumentPart mainPart)
        {
            const int LandscapeWidth = 16840;  // 11.69 inch
            const int LandscapeHeight = 11906; // 8.27 inch

            var sectionProps = mainPart.Document.Body.Elements<SectionProperties>().LastOrDefault();
    
            if (sectionProps == null)
            {
                sectionProps = new SectionProperties();
                mainPart.Document.Body.Append(sectionProps);
            }

            var pageSize = sectionProps.Elements<PageSize>().FirstOrDefault();

            if (pageSize != null)
            {
                pageSize.Orient = PageOrientationValues.Landscape;
                pageSize.Width = (UInt32Value)LandscapeWidth;
                pageSize.Height = (UInt32Value)LandscapeHeight;
            }
            else
            {
                pageSize = new PageSize() 
                { 
                    Orient = PageOrientationValues.Landscape, 
                    Width = (UInt32Value)LandscapeWidth, 
                    Height = (UInt32Value)LandscapeHeight 
                };
                sectionProps.Append(pageSize);
            }
        }

        // ... other methods ...
    }
}
