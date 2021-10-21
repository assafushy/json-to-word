using System.IO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Reflection;
using JsonToWord.Converters;
using JsonToWord.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using log4net;


namespace JsonToWord.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WordController : ControllerBase
    {

        [HttpGet("status")]
        public IActionResult GetStatus()
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);

            return Ok($"{DateTime.Now} Online - Version {versionInfo.FileVersion}");
        }

        [HttpPost("create")]
        public IActionResult CreateWordDocument(dynamic json)
        {
            log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
            try
            {
                string test = json.ToString();
                test = test.Replace("\\\\", "\\");
                var settings = new JsonSerializerSettings();
                settings.Converters.Add(new WordObjectConverter());
                var wordModel = JsonConvert.DeserializeObject<WordModel>(json.ToString(), settings);

                log.Info("Initilized word model object");

                var wordService = new WordService(wordModel);


                var document = wordService.Create();
                log.Info("Created word document");

                return Ok(document);
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

                var wordService = new WordService(wordModel);


                var document = wordService.Create();
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