﻿namespace JsonToWord.Models.S3
{
    public class UploadProperties
    {
        public string BucketName { get; set; }
        public string FileName { get; set; }
        public string LocalFilePath { get; set; }
        public string AwsAccessKeyId { get; set; }
        public string AwsSecretAccessKey { get; set; }
        public string Region { get; set; }


        public UploadProperties(string BucketName, string LocalFilePath, string fileName)
        {
            this.BucketName = BucketName;
            this.LocalFilePath = LocalFilePath;
            this.FileName = fileName;
        }

    }
}
