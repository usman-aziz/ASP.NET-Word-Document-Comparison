using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Word_Comparer.Models;

namespace Word_Comparer.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public FileResult UploadFiles(List<IFormFile> files, string outputFormat)
        {
            if(files.Count()==0)
            {
                return null;
            }
            string fileName = "result.docx";
            // upload files
            var file1 = Path.Combine("wwwroot/uploads", files[0].FileName);
            var file2 = Path.Combine("wwwroot/uploads", files[1].FileName);
            using (var stream = new FileStream(file1, FileMode.Create))
            {
                files[0].CopyTo(stream);
            }
            using (var stream = new FileStream(file2, FileMode.Create))
            {
                files[1].CopyTo(stream);
            }
            // Load Word documents
            Document doc1 = new Document(file1);
            Document doc2 = new Document(file2);

            CompareOptions compareOptions = new CompareOptions();
            compareOptions.IgnoreFormatting = true;
            compareOptions.IgnoreCaseChanges = true;
            compareOptions.IgnoreComments = true;
            compareOptions.IgnoreTables = true;
            compareOptions.IgnoreFields = true;
            compareOptions.IgnoreFootnotes = true;
            compareOptions.IgnoreTextboxes = true;
            compareOptions.IgnoreHeadersAndFooters = true;
            compareOptions.Target = ComparisonTargetType.New;

            var outputStream = new MemoryStream();
            doc1.Compare(doc2, "John Doe", DateTime.Now, compareOptions);
            if (outputFormat == "DOCX")
            {
                doc1.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0;
                // Return generated Word file
                return File(outputStream, System.Net.Mime.MediaTypeNames.Application.Rtf, fileName);
            }
            else
            {
                fileName = "result.pdf";
                doc1.Save(outputStream, SaveFormat.Pdf);
                outputStream.Position = 0;
                // Return generated PDF file
                return File(outputStream, System.Net.Mime.MediaTypeNames.Application.Pdf, fileName);
            }    
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
