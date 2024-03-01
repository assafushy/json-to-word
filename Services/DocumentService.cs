using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
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
                    SetLandscape(mainPart);
                    mainPart.Document.Save();
                }

                File.WriteAllBytes(destinationFile, templateStream.ToArray());
            }

            return destinationFile;
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

        //internal void RunMacro(string documentPath, string macroName, StreamWriter sw)
        //{
        //    Application wordApp = null;
        //    Document wordDoc = null;
        //    var missing = System.Reflection.Missing.Value;
        //    object[] args = new object[1];
        //    args[0] = macroName;
        //    sw.WriteLine("before try in macro");
        //    sw.Flush();
        //    try
        //    {
        //        wordApp = new Application { Visible = false };
        //        sw.WriteLine("befor open document " + documentPath);
        //        sw.Flush();
        //        wordDoc = wordApp.Documents.Open(documentPath, ReadOnly: false, Visible: false);
        //        sw.WriteLine("after open document "+ documentPath);
        //        sw.Flush();
        //        wordApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, 
        //            null, wordApp, args);
        //        wordDoc.Close(true, missing, missing);
        //        wordApp.Quit(true, missing, missing);
        //        sw.WriteLine("end of try");
        //        sw.Flush();
        //    }
        //    catch (Exception exception)
        //    {
        //        sw.WriteLine(exception.Message);
        //        sw.Flush();
        //        //ToDo: write exception to log
        //    }
        //    finally
        //    {
        //        sw.WriteLine("before finaly");
        //        sw.Flush();
        //        if (wordDoc != null)
        //        {
        //            Marshal.FinalReleaseComObject(wordDoc);
        //            wordDoc = null;
        //        }

        //        if (wordApp != null)
        //        {
        //            Marshal.FinalReleaseComObject(wordApp);
        //            wordApp = null;
        //        }
        //        sw.WriteLine("after finaly");
        //        sw.Flush();
        //    }
        //}
    }
}
