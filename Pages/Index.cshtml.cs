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
            var webRootFolder = _hostingEnvironment.WebRootPath;
            var fileName = @"report.xlsx";
            FileInfo file = new FileInfo(Path.Combine(webRootFolder,fileName));
            var spreadSheetName = "Report";
            var startRow = 2;
            var package = ExcelReportHelper.CreateExcelPackage(file);
            var spreedSheet = ExcelReportHelper.CreateWorkSheet(package, spreadSheetName);
            
            string nairaSymbol = ((char)8358).ToString();

            startRow = ExcelReportHelper.AddSealDecription(startRow, new List<SealDescription>() { new SealDescription { AuditTrail="First of july\n brake shore",
            Seal="1006",SortedValue="FEWFEW",DeclearedValue="DWEDF",Client="FCDF",ATM="DEWD"} }, spreedSheet);

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddDailyConsolidatedReport(startRow, new Dictionary<string, CurrencyType>() {

                {@"1,000",new CurrencyType{Mint=1526189000,CAC=28000000,ATM=2000000,CAD=70515000,AE=48852000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1735556000  } },
                {@"500",new CurrencyType{Mint=668100000,CAC=1388150000,ATM=2150000,CAD=34267000,AE=70000000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=2165667000  } },
                {@"200",new CurrencyType{Mint=11420000 ,CAC=320000 ,ATM=0,CAD=6526400,AE=1100000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=19366400  } },
                {@"100",new CurrencyType{Mint=7790000,CAC=20000,ATM=0,CAD=1828800,AE=100000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=9738800  } },
                {@"50",new CurrencyType{Mint=3425000 ,CAC=0,ATM=0,CAD=578550,AE=25000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB= 4028550  } },
                {@"20",new CurrencyType{Mint=906000,CAC=0,ATM=0,CAD=242480,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1148480  } },
                {@"10",new CurrencyType{Mint=4000,CAC=2000,ATM=0,CAD=53510,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=59510  } },
                {@"5",new CurrencyType{Mint=511000,CAC=500,ATM=0,CAD=1875,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=513375  } },
                {@"2",new CurrencyType{Mint=306888,CAC=0,ATM=0,CAD=0,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB= 306888  } },
                {@"1",new CurrencyType{Mint= 505013,CAC=0,ATM=0,CAD=0,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=505013  } },
                {@"50k",new CurrencyType{Mint= 274758,CAC=0,ATM=0,CAD=0,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=274758  } },
                {@"25k",new CurrencyType{Mint=8,CAC=0,ATM=0,CAD=0,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=8  } },
                {@"10k",new CurrencyType{Mint=114,CAC=0,ATM=0,CAD=0,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=114  } },
                {@"1k",new CurrencyType{Mint=10,CAC=0,ATM=0,CAD=0,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=10  } },
                {@"TOTAL",new CurrencyType{Mint= 2219431791,CAC=1416492500,ATM=4150000,CAD=114013615,AE=120077000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=3937164906  } },
                

            }, spreedSheet);

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddDominationProcess(startRow, new Dictionary<string, Domination> {

                {nairaSymbol+"1000", new Domination{ Denomination=nairaSymbol+"1000",Box=4,Counterfeit=0,Shortages=10,Mixup=0,Overages=0,Total=10} },
                {nairaSymbol+"500", new Domination{ Denomination=nairaSymbol+"500",Box=1,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {nairaSymbol+"200", new Domination{ Denomination=nairaSymbol+"200",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {nairaSymbol+"100", new Domination{ Denomination=nairaSymbol+"100",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {nairaSymbol+"50", new Domination{ Denomination=nairaSymbol+"50",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {nairaSymbol+"20", new Domination{ Denomination=nairaSymbol+"20",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {nairaSymbol+"10", new Domination{ Denomination=nairaSymbol+"10",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {nairaSymbol+"5", new Domination{ Denomination=nairaSymbol+"5",Box=5,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {"TOTAL", new Domination{ Denomination=nairaSymbol+"5",Box=5,Counterfeit=0,Shortages=10,Mixup=0,Overages=0,Total=10} },
                {"VALUE "+nairaSymbol, new Domination{ Denomination=nairaSymbol+"5",Box=17990000,Counterfeit=0,Shortages=10000,Mixup=0,Overages=0,Total=10000} },
            }, spreedSheet);

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
