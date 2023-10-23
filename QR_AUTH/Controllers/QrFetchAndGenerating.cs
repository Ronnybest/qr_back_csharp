using Microsoft.AspNetCore.Mvc;
using Net.Codecrete.QrCodeGenerator;
using Newtonsoft.Json;
using QR_AUTH.Models;
using System.Net;
using Xceed.Words.NET;

namespace QR_AUTH.Controllers
{
    internal class FileStreamDelete : FileStream
    {
        readonly string path;

        public FileStreamDelete(string path, FileMode mode) : base(path, mode) // NOTE: must create all the constructors needed first
        {
            this.path = path;
        }

        protected override void Dispose(bool disposing) // NOTE: override the Dispose() method to delete the file after all is said and done
        {
            base.Dispose(disposing);
            if (disposing)
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }
    }

    [Route("api/[controller]")]
    [ApiController]
    public class QrFetchAndGenerating : ControllerBase
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private static List<QrModel.Merchant> _qrData = new List<QrModel.Merchant>();
        private readonly IConfiguration _configuration;
        private readonly DocX _document;
        public QrFetchAndGenerating(IHttpClientFactory httpClientFactory, IConfiguration configuration)
        {
            _httpClientFactory = httpClientFactory;

            _configuration = configuration;
            _document = DocX.Create(@"Assets/WordDosc/qr_codes.docx");
            IConfigurationRoot root = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .AddEnvironmentVariables().Build();
        }

        //private void ClearWordDocument(string fileName)
        //{
        //    using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, true))
        //    {
        //        var firstElement = doc.MainDocumentPart.Document.Body.Elements().FirstOrDefault();
        //        firstElement?.Remove();

        //        foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
        //        {
        //            Header header = headerPart.Header;
        //            foreach (var paragraph in header.Descendants())
        //            {
        //                paragraph.RemoveAllChildren();
        //            }
        //        }

        //        foreach (var footerPart in doc.MainDocumentPart.FooterParts)
        //        {
        //            Footer footer = footerPart.Footer;
        //            foreach (var paragraph in footer.Descendants())
        //            {
        //                paragraph.RemoveAllChildren();
        //            }
        //        }

        //        doc.MainDocumentPart.Document.Save();
        //    }
        //}

        private void GenerateAndSaveQrCode(string qrData, string fileName)
        {
            var qr = QrCode.EncodeText(qrData, QrCode.Ecc.Medium);
            qr.SaveAsPng(fileName, scale: 10, border: 2);
        }

        private void DeleteFilesInFolder(string folderPath)
        {
            try
            {
                string[] files = Directory.GetFiles(folderPath);
                foreach (string filePath in files)
                {
                    System.IO.File.Delete(filePath);
                    // Console.WriteLine($"Удален файл: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при удалении файлов: {ex.Message}");
            }
        }

        private FileContentResult returnFile (string path)
        {
            using (var fileStream = new FileStreamDelete(path, FileMode.Open))
            {
                var memoryStream = new MemoryStream();
                fileStream.CopyTo(memoryStream);

                Response.Headers.Add("Content-Disposition", "attachment; filename=qr_codes.docx");
                Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                DeleteFilesInFolder("Assets/GenerateAllRequest");
                return File(memoryStream.ToArray(), Response.ContentType);
            }
        }
        private void InsertTextAndImage(string text, string imageFileName, string link)
        {
            var p = _document.InsertParagraph();
            p.AppendLine("Название: " + text);
            p.AppendLine("---------------------------------------------------------------------------------");
            var image =_document.AddImage(imageFileName);
            var picture = image.CreatePicture(300, 300);
            p.AppendPicture(picture);
            p.AppendLine("Ссылка: " + link).InsertPageBreakAfterSelf();
            
            //builder.Writeln(text);
            //builder.Writeln("---------------------------------------------------------------------------------");
            //builder.InsertImage(imageFileName);
            //builder.InsertBreak(BreakType.LineBreak);
            //builder.Writeln(link);
            //builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        [HttpGet("get-qr")]
        public async Task<IActionResult> GetQr()
        {
            try
            {
                var handler = new HttpClientHandler
                {
                    UseProxy = true,
                    Proxy = WebRequest.GetSystemWebProxy(), // Использует системные настройки прокси Windows
                    UseDefaultCredentials = true // Использует учетные данные пользователя для прокси, если они настроены в Windows
                };

                var client = new HttpClient(handler);
                client.DefaultRequestHeaders.Add("X-Key", "eyJhbGciOiJSU0EtT0FFUC0yNTYiLCJlbmMiOiJBMjU2Q0JDLUhTNTEyIiwiaWF0IjoxNjYwNjQ4MjgxfQ.LYim6fg00v5ZHgbHUlArTZfG9-IwdEXKSfUySbBAw8EFaRrTsjdo8NKRmsulNpnCOwTQ95MCBrJbiGcJmD3o1xDaL8oujbGndh2ZuYoKxNRn6rNhLlCutg4WgYPUQoYOYAGFwCZ3SQbHgYbUAsVXqp3SqkPqg1dx6MzbgtJKKqTnqFWMNdbodX7mV0Dg2ySD4unLKuWFkbXHvyNU1yCPnyW6vdb4lHIjzTfsFNAU-r-FiNSJepq1x5P_MuUR117bzwUtwqdrBia_1pvqp5zGmLgw2UjXGkN7uQZ-um702_uPKnglBGGHTGcfGxLKQzH_royOAqZagL6cApwqtibLCQ.21ezAEMS2J46jpBKje62-g.mcDdHgYfmOs4s63NNGC_2m7inOYRKMw-HmU1XbwhkSyHrz8p4Yn_7NB6R6xEtirLCc6qW8a90yh11IYOb4R01w._cw0aV7gsGP8NX5e1qCVr2SX_opHEiEiTH3mF1citYY");
                client.DefaultRequestHeaders.Add("X-App-UUID", "d607d6f6-b316-4207-b84b-4802a31ef0f5");
                //client.DefaultRequestHeaders.Add("Connection", "keep-alive");
                //client.DefaultRequestHeaders.Add("User-Agent", "PostmanRuntime/7.34.0");
                var data = await client.GetAsync(_configuration["QrDataGetLink"]);
                if (data.IsSuccessStatusCode)
                {
                    var content = await data.Content.ReadAsStringAsync();
                    QrModel.Root root = JsonConvert.DeserializeObject<QrModel.Root>(content);
                    if (root != null)
                    {
                        _qrData = root.Merchants;
                    }

                    return Ok(content);
                }
                else
                {
                    return StatusCode((int)data.StatusCode, "Request failed");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return StatusCode(500, $"Internal server error: {e.Message}");
            }
        }

        [HttpGet("generate")]
        public IActionResult GenerateQr()
        {
            try
            {
                //DeleteFilesInFolder("Assets/WordDosc");
                foreach (var singleQr in _qrData)
                {
                    string fileName = $"Assets/GenerateAllRequest/{singleQr.Id}.png";
                    // Console.WriteLine(fileName);
                    GenerateAndSaveQrCode(singleQr.qr_data, fileName);
                    InsertTextAndImage(singleQr.Title, fileName, link: singleQr.qr_data);
                }
                _document.Save();
                string wordFileName = "Assets/WordDosc/qr_codes.docx";
                //doc.Save(wordFileName);

                //ClearWordDocument(fileName: wordFileName);
                //FileStreamResult response;
                DeleteFilesInFolder("Assets/GenerateAllRequest");
                return returnFile(wordFileName);

                //var fileStream = System.IO.File.OpenRead(wordFileName);

                //var response = File(fileStream,
                //    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "qr_codes.docx");
                //Response.Headers.Add("Content-Disposition", "attachment; filename=qr_codes.docx");
                //fileStream.Close();
                //DeleteFilesInFolder("Assets/WordDosc");
                //DeleteFilesInFolder("Assets/GenerateAllRequest");
                //return response;
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception);
                return StatusCode(500, $"Internal server error: {exception.Message}");
            }
        }

        [HttpPost("generate-selected")]
        public IActionResult GenerateSelected([FromBody] List<int> idList)
        {
            //DeleteFilesInFolder("Assets/WordDosc");
            //Document doc = new Document();
            //DocumentBuilder builder = new DocumentBuilder(doc);
            foreach (var data in idList)
            {
                Console.WriteLine(data);
                var pointData = _qrData.FirstOrDefault(item => item.Id == data);
                string tempQrCodeFileName = $"Assets/SelectedRequest/{pointData.Id}.png";
                GenerateAndSaveQrCode(pointData.qr_data, tempQrCodeFileName);
                InsertTextAndImage(pointData.Title, tempQrCodeFileName, link: pointData.qr_data);
            }

            string wordFileName = "Assets/WordDosc/qr_codes.docx";
            _document.Save();
            //ClearWordDocument(wordFileName);
            DeleteFilesInFolder("Assets/SelectedRequest");
            return returnFile(wordFileName);
            //var fileStream = System.IO.File.OpenRead(wordFileName);
            //var response = File(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            //    "qr_codes.docx");
            //Response.Headers.Add("Content-Disposition", "attachment; filename=qr_codes.docx");
            //// fileStream.Close();
            //DeleteFilesInFolder("Assets/SelectedRequest");
            //// DeleteFilesInFolder("Assets/WordDosc"); // Возможно, ошибка в имени папки
            //return response;
        }

        [HttpGet("get-by-one")]
        public IActionResult GetByOneId(int id)
        {
            try
            {
                DeleteFilesInFolder("Assets/OneIdRequest");
                // Console.WriteLine(id);
                var pointData = _qrData.FirstOrDefault(item => item.Id == id);
                string tempQrCodeFileName = $"Assets/OneIdRequest/{pointData.Id}.png";
                GenerateAndSaveQrCode(pointData.qr_data, tempQrCodeFileName);
                // Console.WriteLine(tempQrCodeFileName);
                pointData.QrLink = tempQrCodeFileName;
                
                return Ok(pointData);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}