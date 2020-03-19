using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using razorReport.Models;
using razorReport.Helper;

namespace razorReport.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public IndexModel(ILogger<IndexModel> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        public void OnGet()
        {
            
        }

        public IActionResult OnPost(){
            Console.WriteLine("here");
            var webRootFolder = _hostingEnvironment.WebRootPath;
            var fileName = @"report.xlsx";
            FileInfo file = new FileInfo(Path.Combine(webRootFolder,fileName));
            var spreadSheetName = "Report";
            var startRow = 2;
            var package = ExcelReportHelper.CreateExcelPackage(file);
            var spreedSheet = ExcelReportHelper.CreateWorkSheet(package, spreadSheetName);
            startRow = ExcelReportHelper.AddSealDecription(startRow, new List<SealDescription>() { new SealDescription { AuditTrail="First of july\n brake shore",
            Seal="1006",SortedValue="FEWFEW",DeclearedValue="DWEDF",Client="FCDF",ATM="DEWD"} }, spreedSheet);
            package.Save();
            var result = PhysicalFile(Path.Combine(webRootFolder, fileName), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            Response.Headers["Content-Disposition"] = new ContentDispositionHeaderValue("attachment")
            {
                FileName = file.Name
            }.ToString();
            return result;
        }
    }
}
