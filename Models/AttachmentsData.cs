using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace JsonToWord.Models
{
    public class AttachmentsData
    {
        public Uri attachmentMinioPath { get; set; }
        public String minioFileName { get; set; }
    }
}
