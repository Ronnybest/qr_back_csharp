using Aspose.Words;
using Microsoft.AspNetCore.Mvc;
using Net.Codecrete.QrCodeGenerator;
using Newtonsoft.Json;
using QR_AUTH.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System;
using System.Threading.Tasks;
using Document = Aspose.Words.Document;

namespace QR_AUTH.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class QrFetchAndGenerating : ControllerBase
    {
        private readonly IHttpClientFactory _httpClientFactory;
        private static List<QrModel.Merchant> _qrData = new List<QrModel.Merchant>();
        private static Document Doc = new Document();
        private readonly DocumentBuilder _builder;

        public QrFetchAndGenerating(IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;

            if (Doc == null)
            {
                Doc = new Document();
            }

            _builder = new DocumentBuilder(Doc);
        }

        private void ClearWordDocument(string fileName)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, true))
            {
                var firstElement = doc.MainDocumentPart.Document.Body.Elements().FirstOrDefault();
                firstElement?.Remove();

                foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                {
                    Header header = headerPart.Header;
                    foreach (var paragraph in header.Descendants())
                    {
                        paragraph.RemoveAllChildren();
                    }
                }

                foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                {
                    Footer footer = footerPart.Footer;
                    foreach (var paragraph in footer.Descendants())
                    {
                        paragraph.RemoveAllChildren();
                    }
                }

                doc.MainDocumentPart.Document.Save();
            }
        }

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
                    Console.WriteLine($"Удален файл: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при удалении файлов: {ex.Message}");
            }
        }

        private void InsertTextAndImage(string text, string imageFileName, DocumentBuilder builder)
        {
            builder.Writeln(text);
            builder.Writeln("---------------------------------------------------------------------------------");
            builder.InsertImage(imageFileName);
            builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        [HttpGet("get-qr")]
        public async Task<IActionResult> GetQr()
        {
            try
            {
                var client = _httpClientFactory.CreateClient();
                var data = await client.GetAsync("https://localhost:7194/qr-list-get");
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
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                DeleteFilesInFolder("Assets/WordDosc");
                foreach (var singleQr in _qrData)
                {
                    string fileName = $"Assets/{singleQr.Id}.png";
                    // Console.WriteLine(fileName);
                    GenerateAndSaveQrCode(singleQr.qr_data, fileName);
                    InsertTextAndImage(singleQr.Title, fileName, builder: builder);
                }

                string wordFileName = "Assets/WordDosc/qr_codes.docx";
                doc.Save(wordFileName);

                ClearWordDocument(fileName: wordFileName);
                var fileStream = System.IO.File.OpenRead(wordFileName);

                var response = File(fileStream,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "qr_codes.docx");
                Response.Headers.Add("Content-Disposition", "attachment; filename=qr_codes.docx");
                //fileStream.Close();
                DeleteFilesInFolder("Assets/WordDosc");
                //DeleteFilesInFolder("Assets/WordDosc");
                return response;
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
            DeleteFilesInFolder("Assets/WordDosc");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            foreach (var data in idList)
            {
                Console.WriteLine(data);
                var pointData = _qrData.FirstOrDefault(item => item.Id == data);
                string tempQrCodeFileName = $"Assets/SelectedRequest/{pointData.Id}.png";
                GenerateAndSaveQrCode(pointData.qr_data, tempQrCodeFileName);
                InsertTextAndImage(pointData.Title, tempQrCodeFileName, builder);
            }

            string wordFileName = "Assets/WordDosc/qr_codes.docx";
            doc.Save(wordFileName);
            ClearWordDocument(wordFileName);
            var fileStream = System.IO.File.OpenRead(wordFileName);
            var response = File(fileStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "qr_codes.docx");
            Response.Headers.Add("Content-Disposition", "attachment; filename=qr_codes.docx");
            // fileStream.Close();
            DeleteFilesInFolder("Assets/SelectedRequest");
            // DeleteFilesInFolder("Assets/WordDosc"); // Возможно, ошибка в имени папки
            return response;
        }

        [HttpGet("get-by-one")]
        public IActionResult GetByOneId(int id)
        {
            try
            {
                DeleteFilesInFolder("Assets/OneIdRequest");
                Console.WriteLine(id);
                var pointData = _qrData.FirstOrDefault(item => item.Id == id);
                string tempQrCodeFileName = $"Assets/OneIdRequest/{pointData.Id}.png";
                GenerateAndSaveQrCode(pointData.qr_data, tempQrCodeFileName);
                Console.WriteLine(tempQrCodeFileName);
                // _builder.Writeln(pointData.Title);
                // _builder.Writeln("---------------------------------------------------------------------------------");
                // _builder.InsertImage(tempQrCodeFileName);
                // 
                // string currentDirectory = Directory.GetCurrentDirectory();
                // string targetFolder = "Assets/OneIdRequest";
                // string sourceFilePath = Path.Combine(currentDirectory, tempQrCodeFileName);
                // string targetFilePath = Path.Combine(targetFolder, tempQrCodeFileName);
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