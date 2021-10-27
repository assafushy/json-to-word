﻿using System;
using System.Threading.Tasks;
using JsonToWord.Models.S3;


namespace JsonToWord.Services.Interfaces
{
    public interface IAWSS3Service
    {
        AWSUploadResult<string> GenerateAwsFileUrl(string bucketName, string key, string region, bool useRegion = true);
        Task<AWSUploadResult<string>> UploadFileToS3BucketAsync(UploadProperties uploadProperties);
        string DownloadFileFromS3BucketAsync(Uri webPath, string filename);
        void CleanUp(string filename);

    }
}
