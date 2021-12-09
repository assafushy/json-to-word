using JsonToWord.Converters;
using JsonToWord.Models;
using JsonToWord.Models.S3;
using JsonToWord.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading.Tasks;

namespace JsonToWord.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WordController : ControllerBase
    {
        private readonly IAWSS3Service _aWSS3Service;
        private readonly IWordService _wordService;

        public WordController(IAWSS3Service aWSS3Service,IWordService wordService)
        { 
            _aWSS3Service = aWSS3Service;
            _wordService = wordService;
        }

        [HttpGet("status")]
        public IActionResult GetStatus()
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);

            return Ok($"{DateTime.Now} Online - Version {versionInfo.FileVersion}");
        }

        [HttpPost("create")]
        public async Task<IActionResult> CreateWordDocument(dynamic json)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            try
            {
                var settings = new JsonSerializerSettings();
                settings.Converters.Add(new WordObjectConverter());
                WordModel wordModel = JsonConvert.DeserializeObject<WordModel>(json.ToString(), settings);
                string fullpath = _aWSS3Service.DownloadFileFromS3BucketAsync(wordModel.TemplatePath,wordModel.UploadProperties.FileName);
                wordModel.LocalPath = fullpath;
                log.Info("Initilized word model object");

                //var wordService = new WordService();
                var documentPath = _wordService.Create(wordModel);
                log.Info("Created word document");

                _aWSS3Service.CleanUp(fullpath);

                wordModel.UploadProperties.LocalFilePath = documentPath;

                AWSUploadResult<string> Response = await _aWSS3Service.UploadFileToMinioBucketAsync(wordModel.UploadProperties);

                _aWSS3Service.CleanUp(documentPath);

                if (Response.Status)
                {
                    return Ok(Response.Data);
                }
                else
                {
                    return StatusCode(Response.StatusCode);
                }

            }
            catch (Exception e)
            {
                string logPath = @"c:\logs\prod\JsonToWord.log";
                System.IO.File.AppendAllText(logPath, string.Format("\n{0} - {1}", DateTime.Now, e));
                throw;
            }
        }

        [HttpPost("create-by-file")]
        public IActionResult CreateWordDocumentByFile(dynamic json)
        {
            try
            {
                string file = json.jsonFilePath;
                string text = System.IO.File.ReadAllText(file);
                json = JObject.Parse(text);

                log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

                string test = json.ToString();
                test = test.Replace("\\\\", "\\");
                var settings = new JsonSerializerSettings();
                settings.Converters.Add(new WordObjectConverter());
                var wordModel = JsonConvert.DeserializeObject<WordModel>(json.ToString(), settings);

                log.Info("Initilized word model object");

                var wordService = new WordService();


                var document = wordService.Create(wordModel);
                log.Info("Created word document");

                return Ok(document);
            }
            catch (Exception e)
            {
                return null;
            }

        }
    }
}