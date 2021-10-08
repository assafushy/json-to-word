using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Web.Http;
using JsonToWord.Converters;
using JsonToWord.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using log4net;

namespace JsonToWord.Controllers
{
    public class WordController : ApiController
    {
        [HttpGet]
        [ActionName("status")]
        public HttpResponseMessage GetStatus()
        {
            var versionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);

            var response = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent($"{DateTime.Now} Online - Version {versionInfo.FileVersion}")
            };

            return response;
        }

        [HttpPost]
        [ActionName("create")]
        public HttpResponseMessage CreateWordDocument(dynamic json)
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

                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(document)
                };

                return response;
            }
            catch (Exception e)
            {
               
                string logPath = @"c:\logs\prod\JsonToWord.log";
                File.AppendAllText(logPath, string.Format("\n{0} - {1}", DateTime.Now,e));

                throw;
            }
        }
        [HttpPost]
        [ActionName("create-by-file")]
        public HttpResponseMessage CreateWordDocumentByFile(dynamic json)
        {
            try
            {
                string file = json.jsonFilePath;
                string text = File.ReadAllText(file);
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

                var response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(document)
                };
                return response;
            }
            catch(Exception e) {
                return null;
            }

            
        }

    }
}
