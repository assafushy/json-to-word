using JsonToWord.Models.S3;
using System;
using System.Collections.Generic;

namespace JsonToWord.Models
{
    public class WordModel
    {
        public UploadProperties UploadProperties { get; set; }
        public Uri TemplatePath { get; set; }
        public AttachmentsData[] MinioAttachmentData { get; set; }
        public List<WordContentControl> ContentControls { get; set; }
        public string LocalPath { get; set; }
        public List<JsonData> JsonDataList { get; set; }
    }
}