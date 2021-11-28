using System;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Xml;
using Microsoft.Win32;
//var wordApplication = new Microsoft.Office.Interop.Word.Application();


namespace JsonToWord.Models
{
    public class OpenXmlEmbeddedObject
    {
        private Icon _objectIcon;
        private string _objectIconFile;
        private string _oleImageStyle;
        private static FileInfo _fileInfo;
        private static string _filePathAndName;
        private static bool _displayAsIcon;
        private static bool _objectIsPicture;
        private object _objectMissing = System.Reflection.Missing.Value;
        private object _objectFalse = false;
        private object _objectTrue = true;
        private string _fileType;
        private string _fileContentType;

        private const string DefaultOleContentType = "application/vnd.openxmlformats-officedocument.oleObject";
        private const string CsvOleContentType = "application/vnd.ms-excel.sheet.macroEnabled.12";
        private const string OleObjectDataTag = "application/vnd";
        private const string OleImageDataTag = "image/x-emf";

        public string FileType
        {
            get
            {
                if (string.IsNullOrEmpty(_fileType) && _fileInfo != null)
                {
                   // _fileType = GetFileType(_fileInfo, false);
                }

                return _fileType;
            }
        }

        public string FileContentType
        {
            get
            {
                if (string.IsNullOrEmpty(_fileContentType) && _fileInfo != null)
                {
                    //_fileContentType = GetFileContentType(_fileInfo);

                    if (_fileInfo.Extension == ".csv" && _fileContentType.StartsWith("application/vnd.ms-excel"))
                    {
                        _fileContentType = CsvOleContentType;
                    }
                    else if (!_fileContentType.Contains("officedocument"))
                    {
                        _fileContentType = DefaultOleContentType;
                    }
                }

                return _fileContentType;
            }
        }

        //public static string GetFileContentType(FileInfo fileInfo)
        //{
        //    if (fileInfo == null)
        //    {
        //        throw new ArgumentNullException(nameof(fileInfo));
        //    }

        //    var mime = "application/octetstream";

        //    var ext = Path.GetExtension(fileInfo.Name).ToLower();

        //    var rk = Registry.ClassesRoot.OpenSubKey(ext);

        //    if (rk?.GetValue("Content Type") != null)
        //    {
        //        mime = rk.GetValue("Content Type").ToString();
        //    }

        //    return mime;
        //}

        public bool ObjectIsOfficeDocument => FileContentType != DefaultOleContentType;

        public bool ObjectIsPicture => _objectIsPicture;

        public string OleObjectBinaryData { get; set; }

        public string OleImageBinaryData { get; set; }

        /// <summary>
        /// The OpenXml information for the Word Application that is created (Make-Shoft Code Reflector)
        /// </summary>
        public string WordOpenXml { get; set; }

        /// <summary>
        /// The XmlDocument that is created based on the OpenXml Data from WordOpenXml
        /// </summary>
        public XmlDocument OpenXmlDocument
        {
            get
            {
                if (_openXmlDocument == null && !string.IsNullOrEmpty(WordOpenXml))
                {
                    _openXmlDocument = new XmlDocument();
                    _openXmlDocument.LoadXml(WordOpenXml);
                }

                return _openXmlDocument;
            }
        }
        private XmlDocument _openXmlDocument;

        /// <summary>
        /// The XmlNodeList, for all Nodes containing 'binaryData'
        /// </summary>
        public XmlNodeList BinaryDataXmlNodesList
        {
            get
            {
                if (_binaryDataXmlNodesList == null && OpenXmlDocument != null)
                {
                    _binaryDataXmlNodesList = OpenXmlDocument.GetElementsByTagName("pkg:binaryData");
                }

                return _binaryDataXmlNodesList;
            }
        }
        private XmlNodeList _binaryDataXmlNodesList;

        public Icon ObjectIcon => _objectIcon ?? (_objectIcon = Icon.ExtractAssociatedIcon(_filePathAndName));

        public string ObjectIconFile
        {
            get
            {
                if (string.IsNullOrEmpty(_objectIconFile))
                {
                    _objectIconFile = $"{_filePathAndName.Replace(".", "")}.ico";
                }

                return _objectIconFile;
            }
        }

        public string OleImageStyle
        {
            get
            {
                if (string.IsNullOrEmpty(_oleImageStyle) && !string.IsNullOrEmpty(WordOpenXml))
                {
                    XmlNodeList xmlNodeList = OpenXmlDocument.GetElementsByTagName("v:shape");
                    if (xmlNodeList.Count > 0)
                    {
                        var xmlAttributeCollection = xmlNodeList[0].Attributes;
                        if (xmlAttributeCollection != null)
                            foreach (XmlAttribute attribute in xmlAttributeCollection)
                            {
                                if (attribute.Name == "style")
                                {
                                    _oleImageStyle = attribute.Value;
                                }
                            }
                    }
                }

                return _oleImageStyle;
            }

            set => _oleImageStyle = value;
        }

        #region Constructor

        /// <summary>
        /// Generates binary information for the file being passed in
        /// </summary>
        /// <param name="fileInfo">The FileInfo object for the file to be embedded</param>
        /// <param name="displayAsIcon">Whether or not to display the file as an Icon (Otherwise it will show a snapshot view of the file)</param>

        public OpenXmlEmbeddedObject(FileInfo fileInfo, bool displayAsIcon)
        {
            _fileInfo = fileInfo;
            _filePathAndName = fileInfo.ToString();
            _displayAsIcon = displayAsIcon;

            SetupOleFileInformation();
        }

        private void SetupOleFileInformation()
        {
           // var wordApplication = new Microsoft.Office.Interop.Word.Application();

            //var wordDocument = wordApplication.Documents.Add(ref _objectMissing, ref _objectMissing,
               // ref _objectMissing, ref _objectMissing);

            object iconObjectFileName = _objectMissing;
            object objectClassType = FileType;
            object objectFilename = _fileInfo.ToString();

            if (_displayAsIcon)
            {
                if (ObjectIcon != null)
                {
                    using (var iconStream = new FileStream(ObjectIconFile, FileMode.Create))
                    {
                        ObjectIcon.Save(iconStream);
                        iconObjectFileName = ObjectIconFile;
                    }
                }

                object objectIconLabel = _fileInfo.Name;

                Thread.Sleep(TimeSpan.FromSeconds(1));

              //  wordDocument.InlineShapes.AddOLEObject(ref objectClassType,
                   // ref objectFilename, ref _objectFalse, ref _objectTrue, ref iconObjectFileName,
                   // ref _objectMissing, ref objectIconLabel, ref _objectMissing);
            }
            else
            {
                try
                {
                    var image = Image.FromFile(_fileInfo.ToString());
                    _objectIsPicture = true;
                    OleImageStyle = $"height:{image.Height}pt;width:{image.Width}pt";

                    //wordDocument.InlineShapes.AddPicture(_fileInfo.ToString(), ref _objectMissing, ref _objectTrue, ref _objectMissing);
                }
                catch
                {
                   // wordDocument.InlineShapes.AddOLEObject(ref objectClassType,
                    //    ref objectFilename, ref _objectFalse, ref _objectFalse, ref _objectMissing, ref _objectMissing,
                     //  ref _objectMissing, ref _objectMissing);
                }
            }

           // WordOpenXml = wordDocument.Range(ref _objectMissing, ref _objectMissing).WordOpenXML;

            if (_objectIsPicture)
            {
                OleObjectBinaryData = GetPictureBinaryData();
                OleImageBinaryData = GetPictureBinaryData();
            }
            else
            {
                OleObjectBinaryData = GetOleBinaryData(OleObjectDataTag);
                OleImageBinaryData = GetOleBinaryData(OleImageDataTag);
            }

            // Not sure why, but Excel seems to hang in the processes if you attach an Excel file…
            // This kills the excel process that has been started < 15 seconds ago (so not to kill the user's other Excel processes that may be open)
            /*if (FileType.StartsWith("Excel"))
            {
                var processes = Process.GetProcessesByName("EXCEL");
                foreach (var process in processes)
                {
                    if (DateTime.Now.Subtract(process.StartTime).Seconds <= 15)
                    {
                        process.Kill();
                        break;
                    }
                }
            }*/

            //wordDocument.Close(ref _objectFalse, ref _objectMissing, ref _objectMissing);
            //wordApplication.Quit(ref _objectMissing, ref _objectMissing, ref _objectMissing);
        }

        private string GetOleBinaryData(string binaryDataXmlTag)
        {
            string binaryData = null;
            if (BinaryDataXmlNodesList != null)
            {
                foreach (XmlNode xmlNode in BinaryDataXmlNodesList)
                {
                    if (xmlNode.ParentNode?.Attributes != null)
                    {
                        foreach (XmlAttribute attr in xmlNode.ParentNode.Attributes)
                        {
                            if (string.IsNullOrEmpty(binaryData) && attr.Value.Contains(binaryDataXmlTag))
                            {
                                binaryData = xmlNode.InnerText;
                                break;
                            }
                        }
                    }
                }
            }

            return binaryData;
        }

        private string GetPictureBinaryData()
        {
            string binaryData = null;
            if (BinaryDataXmlNodesList != null)
            {
                foreach (XmlNode xmlNode in BinaryDataXmlNodesList)
                {
                    binaryData = xmlNode.InnerText;
                    break;
                }
            }

            return binaryData;
        }

        //public static string GetFileType(FileInfo fileInfo, bool returnDescription)
        //{
        //    if (fileInfo == null)
        //    {
        //        throw new ArgumentNullException(nameof(fileInfo));
        //    }

        //    string description = "File";
        //    if (string.IsNullOrEmpty(fileInfo.Extension))
        //    {
        //        return description;
        //    }
        //    description = $"{fileInfo.Extension.Substring(1).ToUpper()} File";
        //    var typeKey = Registry.ClassesRoot.OpenSubKey(fileInfo.Extension);

        //    if (typeKey == null)
        //        return description;

        //    var type = Convert.ToString(typeKey.GetValue(string.Empty));
        //    var key = Registry.ClassesRoot.OpenSubKey(type);

        //    if (key == null)
        //        return description;

        //    if (returnDescription)
        //    {
        //        description = Convert.ToString(key.GetValue(string.Empty));
        //        return description;
        //    }

        //    return type;
        //}

        #endregion Methods
    }
}