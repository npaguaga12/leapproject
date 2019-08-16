using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;
using DocumentFormat.OpenXml.Packaging;

namespace BackEnd.Controllers
{
    [Route("api")]
    [ApiController]
    public class UploadDownloadController : ControllerBase
    {
        private IHostingEnvironment _hostingEnvironment;
        public UploadDownloadController(IHostingEnvironment environment)
        {
            _hostingEnvironment = environment;
        }


        [HttpPost]
        [Route("upload")]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            var uploads = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");
            //var watermark = Path.Combine(_hostingEnvironment.WebRootPath, "watermarkuploads");
            if (!Directory.Exists(uploads))
            {
                Directory.CreateDirectory(uploads);
            }
            if (file.Length > 0)
            {
                var path = Path.Combine(uploads, file.FileName);

                using (var fileStream = new FileStream(path, FileMode.Create))
                {
                    await file.CopyToAsync(fileStream);
                }
            }
            return Ok();
        }

        [HttpGet]
        [Route("download")]
        public async Task<IActionResult> Download([FromQuery] string file)
        {
            var uploads = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");
            var filePath = Path.Combine(uploads, file);
            if (!System.IO.File.Exists(filePath))
                return NotFound();

            var memory = new MemoryStream();
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
            }
            memory.Position = 0;

            return File(memory, GetContentType(filePath), file);
        }
        //watermark
        [HttpGet]
        [Route("watermarkdownload")]
        public void WatermarkDownload([FromQuery] string file)
        {
            var uploads = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");
            var filepath = Path.Combine(uploads, file);
            var extension = Path.GetExtension(filepath).ToUpper(); //get file extension
            byte[] sourceBytes = System.IO.File.ReadAllBytes(filepath);
            string watermarkedFile;

            using (MemoryStream inMemoryStream = new MemoryStream())
            {
                inMemoryStream.Write(sourceBytes, 0, (int)sourceBytes.Length);

                if (extension == ".DOC" || extension == ".DOCX")
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Open(inMemoryStream, true))
                    {
                        DocWatermark.AddWaterMark(doc);
                        doc.MainDocumentPart.Document.Save();

                        doc.Close();
                        doc.Dispose();
                    }
                    watermarkedFile = @"C:\Users\v-napagu\Desktop\LEAPFile\BackEnd\BackEnd\wwwroot\watermarkuploads\Watermarked_DOC.docx";
                    System.IO.File.WriteAllBytes(watermarkedFile, inMemoryStream.ToArray());
                }

                else if (extension == ".PPT" || extension == ".PPTX") //if the file is a powerpoint use the following watermark method
                {
                    using (PresentationDocument presentationDoc = PresentationDocument.Open(inMemoryStream, true))
                    {
                        SlideMasterPart slideMasterPart1 = presentationDoc.PresentationPart.SlideMasterParts.First();

                        PPTWatermark.ChangeSlideMasterPart(slideMasterPart1);
                    }
                    watermarkedFile = @"C:\Users\v-napagu\Desktop\LEAPFile\BackEnd\BackEnd\wwwroot\watermarkuploads\Watermarked_PPT.pptx";
                    System.IO.File.WriteAllBytes(watermarkedFile, inMemoryStream.ToArray());
                }

                else if (extension == ".XLS" || extension == ".XLSX")
                {
                    using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(inMemoryStream, true))
                    {
                        XLSWatermark.WatermarkAllSheets(spreadsheet);
                    }
                    watermarkedFile = @"C:\Users\v-napagu\Desktop\LEAPFile\BackEnd\BackEnd\wwwroot\watermarkuploads\Watermarked_XLS.xlsx";
                    System.IO.File.WriteAllBytes(watermarkedFile, inMemoryStream.ToArray());
                }
                inMemoryStream.Close();
            }
        }

        //get files that have been uploaded and display them in file list
        [HttpGet]
        [Route("files")]
        public IActionResult Files()
        {
            var result = new List<string>();
            var uploads = Path.Combine(_hostingEnvironment.WebRootPath, "uploads");
            if (Directory.Exists(uploads))
            {
                var provider = _hostingEnvironment.ContentRootFileProvider;
                foreach (string fileName in Directory.GetFiles(uploads))
                {
                    var fileInfo = provider.GetFileInfo(fileName);
                    result.Add(fileInfo.Name);
                }
            }
            return Ok(result);
        }
        //GET watermark files
        [HttpGet]
        [Route("watermarkfiles")]
        public IActionResult WatermarkFiles()
        {
            var result = new List<string>();
            var watermarkFiles = Path.Combine(_hostingEnvironment.WebRootPath, "watermarkuploads");
            if (Directory.Exists(watermarkFiles))
            {
                var provider = _hostingEnvironment.ContentRootFileProvider;
                foreach (string fileName in Directory.GetFiles(watermarkFiles))
                {
                    var fileInfo = provider.GetFileInfo(fileName);
                    result.Add(fileInfo.Name);
                }
            }
            return Ok(result);
        }

        private string GetContentType(string path)
        {
            var provider = new FileExtensionContentTypeProvider();
            string contentType;
            if (!provider.TryGetContentType(path, out contentType))
            {
                contentType = "application/octet-stream";
            }
            return contentType;
        }
    }
}