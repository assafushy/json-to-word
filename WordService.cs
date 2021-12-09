using System;
using DocumentFormat.OpenXml.Packaging;
using JsonToWord.Models;
using JsonToWord.Services;
using System.IO;
using JsonToWord.Services.Interfaces;

namespace JsonToWord
{
    public class WordService : IWordService
    {
        private readonly ContentControlService _contentControlService;
        private readonly FileService _fileService;
        private readonly HtmlService _htmlService;
        private readonly PictureService _pictureService;
        private readonly TableService _tableService;
        private readonly TextService _textService;
        private readonly DocumentService _documentService;

        public WordService()
        {
            _contentControlService = new ContentControlService();
            _fileService = new FileService();
            _htmlService = new HtmlService();
            _pictureService = new PictureService();
            _tableService = new TableService();
            _textService = new TextService();
            _documentService = new DocumentService();
        }

        public string Create(WordModel _wordModel)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

            var documentPath = _documentService.CreateDocument(_wordModel.LocalPath);

            using (var document = WordprocessingDocument.Open(documentPath, true))
            {
                log.Info("Starting on doc path: " + documentPath);

                foreach (var contentControl in _wordModel.ContentControls)
                {
                    _contentControlService.ClearContentControl(document, contentControl.Title, contentControl.ForceClean);

                    foreach (var wordObject in contentControl.WordObjects)
                    {
                        switch (wordObject.Type)
                        {
                            case WordObjectType.File:
                                _fileService.Insert(document, contentControl.Title, (WordAttachment)wordObject);
                                break;
                            case WordObjectType.Html:
                                _htmlService.Insert(document, contentControl.Title, (WordHtml)wordObject);
                                break;
                            case WordObjectType.Picture:
                                _pictureService.Insert(document, contentControl.Title, (WordAttachment)wordObject);
                                break;
                            case WordObjectType.Paragraph:
                                _textService.Write(document, contentControl.Title, (WordParagraph)wordObject);
                                break;
                            case WordObjectType.Table:
                                _tableService.Insert(document, contentControl.Title, (WordTable)wordObject);
                                break;
                            default:
                                throw new ArgumentOutOfRangeException();
                        }
                    }

                    _contentControlService.RemoveContentControl(document, contentControl.Title);
                }
                log.Info("Finished on doc path: " + documentPath);


            }

            //documentService.RunMacro(documentPath, "updateTableOfContent",sw);
            //log.Info("Ran Macro");

            return documentPath;
        }
    }
}