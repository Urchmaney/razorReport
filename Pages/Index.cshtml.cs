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

        #region 
        //An instance of the cashInStocDec class
        private CashInStockDec ss = new CashInStockDec {
            ax1000 = 100993993,
            ax500 = 29043,
            ax200 = 34434,
            ax100 = 33231,
            ax50  = 344343,
            ax20 = 3434,
            ax10 = 9089,
            ax5 = 43434,
            ax2 = 12323,
            ax1 = 23982,
            ax1coin = 12123,
            ax0_5 = 1212,
            ax0_25 = 32323,
            ax0_2 =32323,
            ax0_1 = 3223,

            a1000 = 1000,
            a500 = 29043,
            a200 = 34434,
            a100 = 33231,
            a50  = 344343,
            a20 = 3434,
            a10 = 9089,
            a5 = 43434,
            a2 = 12323,
            a1 = 23982,
            a1coin = 12123,
            a0_5 = 1212,
            a0_25 = 32323,
            a0_2 =32323,
            a0_1 = 3223,

            f1000 = 1000,
            f500 = 29043,
            f200 = 34434,
            f100 = 33231,
            f50  = 344343,
            f20 = 3434,
            f10 = 9089,
            f5 = 43434,
            f2 = 12323,
            f1 = 23982,
            f1coin = 12123,
            f0_5 = 1212,
            f0_25 = 32323,
            f0_2 =32323,
            f0_1 = 3223,


            u1000 = 1000,
            u500 = 29043,
            u200 = 34434,
            u100 = 33231,
            u50  = 344343,
            u20 = 3434,
            u10 = 9089,
            u5 = 43434,
            u2 = 12323,
            u1 = 23982,
            u1coin = 12123,
            u0_5 = 1212,
            u0_25 = 32323,
            u0_2 =32323,
            u0_1 = 3223,

            m1000 = 1000,
            m500 = 29043,
            m200 = 34434,
            m100 = 33231,
            m50  = 344343,
            m20 = 3434,
            m10 = 9089,
            m5 = 43434,
            m2 = 12323,
            m1 = 23982,
            m1coin = 12123,
            m0_5 = 1212,
            m0_25 = 32323,
            m0_2 =32323,
            m0_1 = 3223,
            
        };
        #endregion

        [BindProperty]
        public string ReportType { get; set; }

        public void OnGet()
        {
            
        }

        public IActionResult OnPost() {            
            var webRootFolder = _hostingEnvironment.WebRootPath;
            FileInfo file;
            switch (ReportType)
            {
                case "Broker":
                    file = ExcelReportHelper.GenerateBrokerReport(webRootFolder, ss);
                    break;
                case "Consolidated":
                    file = ExcelReportHelper.GenerateConsolidatedReport(webRootFolder);
                    break;
                case "Daily-Transaction":
                    file = ExcelReportHelper.GenerateDailyTransactionReport(webRootFolder);
                    break;
                default:
                    file = ExcelReportHelper.GenerateOtherReport(webRootFolder);
                    break;
            }
            var result = PhysicalFile(Path.Combine(webRootFolder, file.Name), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            Response.Headers["Content-Disposition"] = new ContentDispositionHeaderValue("attachment")
            {
                FileName = file.Name
            }.ToString();
            return result;
        }
    }
}
