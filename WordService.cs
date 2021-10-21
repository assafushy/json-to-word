using System;
using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using JsonToWord.Services;
using System.IO;

namespace JsonToWord
{
    public class WordService
    {
        private readonly WordModel _wordModel;

        public WordService(WordModel wordModel)
        {
            _wordModel = wordModel;
        }

        public string Create()
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

            var documentService = new DocumentService();

            var documentPath = documentService.CreateDocument(_wordModel.TemplatePath);

            var contentControlService = new ContentControlService();
            var fileService = new FileService();
            var htmlService = new HtmlService();
            var pictureService = new PictureService();
            var tableService = new TableService();
            var textService = new TextService();

            using (var document = WordprocessingDocument.Open(documentPath, true))
            {
                log.Info("Starting on doc path: " + documentPath);

                foreach (var contentControl in _wordModel.ContentControls)
                {
                    contentControlService.ClearContentControl(document, contentControl.Title, contentControl.ForceClean);

                    foreach (var wordObject in contentControl.WordObjects)
                    {
                        switch (wordObject.Type)
                        {
                            //case WordObjectType.File:
                            //    fileService.Insert(document, contentControl.Title, (WordAttachment)wordObject);
                            //    break;
                            case WordObjectType.Html:
                                htmlService.Insert(document, contentControl.Title, (WordHtml)wordObject);
                                break;
                            case WordObjectType.Picture:
                                pictureService.Insert(document, contentControl.Title, (WordAttachment)wordObject);
                                break;
                            case WordObjectType.Paragraph:
                                textService.Write(document, contentControl.Title, (WordParagraph)wordObject);
                                break;
                            //case WordObjectType.Table:
                            //    tableService.Insert(document, contentControl.Title, (WordTable)wordObject);
                            //    break;
                            default:
                                throw new ArgumentOutOfRangeException();
                        }
                    }

                    contentControlService.RemoveContentControl(document, contentControl.Title);
                }
                log.Info("Finished on doc path: " + documentPath);


            }

            //documentService.RunMacro(documentPath, "updateTableOfContent",sw);
            log.Info("Ran Macro");

            return documentPath;
        }
    }
}