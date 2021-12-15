namespace JsonToWord.Models.S3
{
    public class UploadProperties
    {
        public string BucketName { get; set; }
        public string SubDirectoryInBucket { get; set; }
        public string FileName { get; set; }
        public string LocalFilePath { get; set; }
        public string AwsAccessKeyId { get; set; }
        public string AwsSecretAccessKey { get; set; }
        public string Region { get; set; }
        public string ServiceUrl { get; set; }
    }
}
