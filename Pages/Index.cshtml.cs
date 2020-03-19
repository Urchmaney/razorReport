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

        private CashInStockDec ss = new CashInStockDec {
            ax1000 = 1000,
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

        private FileInfo GenerateBrokerReport() {
            var fileName = @"broker-report.xlsx";
            FileInfo file = new FileInfo(Path.Combine(_hostingEnvironment.WebRootPath, fileName));
            var package = ExcelReportHelper.CreateExcelPackage(file);
            var spreedSheet = ExcelReportHelper.CreateWorkSheet(package, "Report");
            var startRow = 2;
            startRow = ExcelReportHelper.AddBankBrokerReport(startRow, "GUARANTY TRUST BANK PLC ILUPEJU CASH CENTER", new Dictionary<string, CurrencyType>() {
                {@"1,000",new CurrencyType{Mint=1526189000,CAC=28000000,ATM=2000000,CAD=70515000,AE=48852000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1735556000  } },
                {@"500",new CurrencyType{Mint=668100000,CAC=1388150000,ATM=2150000,CAD=34267000,AE=70000000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=2165667000  } },
                {@"200",new CurrencyType{Mint=11420000 ,CAC=320000 ,ATM=0,CAD=6526400,AE=1100000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=19366400  } },
                {@"100",new CurrencyType{Mint=7790000,CAC=20000,ATM=0,CAD=1828800,AE=100000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=9738800  } },
                {@"50",new CurrencyType{Mint=3425000 ,CAC=0,ATM=0,CAD=578550,AE=25000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB= 4028550  } },
                {@"20",new CurrencyType{Mint=906000,CAC=0,ATM=0,CAD=242480,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1148480  } },
                {@"10",new CurrencyType{Mint=4000,CAC=2000,ATM=0,CAD=53510,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=59510  } },
                {@"5",new CurrencyType{Mint=511000,CAC=500,ATM=0,CAD=1875,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=513375  } },
                {@"2",null },
                {@"1", null },
                {@"50k", null },
                {@"25k", null },
                {@"10k", null},
                {@"1k", null},
                {@"TOTAL",new CurrencyType{Mint= 2219431791,CAC=1416492500,ATM=4150000,CAD=114013615,AE=120077000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=3937164906  } }
                }, spreedSheet
            ); 

            startRow = ExcelReportHelper.AddSpace(startRow, 1, spreedSheet);
            startRow = ExcelReportHelper.AddBankBrokerReport(startRow, "GUARANTY TRUST BANK PLC ISLAND CASH CENTER", new Dictionary<string, CurrencyType>() {
                {@"1,000",new CurrencyType{Mint=1526189000,CAC=28000000,ATM=2000000,CAD=70515000,AE=48852000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1735556000  } },
                {@"500",new CurrencyType{Mint=668100000,CAC=1388150000,ATM=2150000,CAD=34267000,AE=70000000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=2165667000  } },
                {@"200",new CurrencyType{Mint=11420000 ,CAC=320000 ,ATM=0,CAD=6526400,AE=1100000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=19366400  } },
                {@"100",new CurrencyType{Mint=7790000,CAC=20000,ATM=0,CAD=1828800,AE=100000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=9738800  } },
                {@"50",new CurrencyType{Mint=3425000 ,CAC=0,ATM=0,CAD=578550,AE=25000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB= 4028550  } },
                {@"20",new CurrencyType{Mint=906000,CAC=0,ATM=0,CAD=242480,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1148480  } },
                {@"10",new CurrencyType{Mint=4000,CAC=2000,ATM=0,CAD=53510,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=59510  } },
                {@"5",new CurrencyType{Mint=511000,CAC=500,ATM=0,CAD=1875,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=513375  } },
                {@"2",null },
                {@"1", null },
                {@"50k", null },
                {@"25k", null },
                {@"10k", null},
                {@"1k", null},
                {@"TOTAL",new CurrencyType{Mint= 2219431791,CAC=1416492500,ATM=4150000,CAD=114013615,AE=120077000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=3937164906  } }
                }, spreedSheet
            );  

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddBankBrokerReport(startRow, "GUARANTY TRUST BANK PLC BOTH CASH CENTER", new Dictionary<string, CurrencyType>() {
                {@"1,000",new CurrencyType{Mint=1526189000,CAC=28000000,ATM=2000000,CAD=70515000,AE=48852000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1735556000  } },
                {@"500",new CurrencyType{Mint=668100000,CAC=1388150000,ATM=2150000,CAD=34267000,AE=70000000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=2165667000  } },
                {@"200",new CurrencyType{Mint=11420000 ,CAC=320000 ,ATM=0,CAD=6526400,AE=1100000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=19366400  } },
                {@"100",new CurrencyType{Mint=7790000,CAC=20000,ATM=0,CAD=1828800,AE=100000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=9738800  } },
                {@"50",new CurrencyType{Mint=3425000 ,CAC=0,ATM=0,CAD=578550,AE=25000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB= 4028550  } },
                {@"20",new CurrencyType{Mint=906000,CAC=0,ATM=0,CAD=242480,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=1148480  } },
                {@"10",new CurrencyType{Mint=4000,CAC=2000,ATM=0,CAD=53510,AE=0,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=59510  } },
                {@"5",new CurrencyType{Mint=511000,CAC=500,ATM=0,CAD=1875,AE=48852000,Today=0,CITI=0,Fidelity=0,DIAMOND=0,IBADANSWAP=0,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=513375  } },
                {@"2",null },
                {@"1", null },
                {@"50k", null },
                {@"25k", null },
                {@"10k", null},
                {@"1k", null},
                {@"TOTAL",new CurrencyType{Mint= 2219431791,CAC=1416492500,ATM=4150000,CAD=114013615,AE=120077000,Today=0,CITI=0,Fidelity=60000000,DIAMOND=0,IBADANSWAP=3000000,KANOSWAP=0,FCMB=0,UBA=0,ATCOB=3937164906  } }
                }, spreedSheet
            );

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddBankBrokerReport(startRow, "GUARANTY TRUST BANK PLC BOTH CASH CENTER", DataHelper.ConvertData(ss), spreedSheet
            );
            package.Save();
            return file;
        }

        private FileInfo GenerateDailyTransactionReport(){
            var fileName = @"daily-trans-report.xlsx";
            FileInfo file = new FileInfo(Path.Combine(_hostingEnvironment.WebRootPath, fileName));
            var package = ExcelReportHelper.CreateExcelPackage(file);
            var spreedSheet = ExcelReportHelper.CreateWorkSheet(package, "Report");
            var startRow = 2;
            startRow = ExcelReportHelper.AddDailyTransaction(startRow, new Dictionary<string, double>{
                {@"#1000",599679000  },
                {@"#500",502852000  },
                {@"#200",3505400},
                {@"#100",2406800},
                {@"#50",259250},
                {@"#20",501540},
                {@"#10",43960 },
                {@"#5",48510  },
                {@"#2",0  },
                {@"#1",0  },
                {@"50k",0  },
                {@"25k",0  },
                {@"10k",0  },
                {@"1k",0  },
                {@"TOTAL",1109296460  }
            },new List<NamedCashDenomnation>{
                new NamedCashDenomnation{ Name="OYINGBO",
                    CashDenomination=new Dictionary<string, double>{
                        { @"#1000",0  },
                        { @"#500",0  },
                        { @"#200",0},
                        { @"#100",0},
                        { @"#50",0},
                        { @"#20",0},
                        { @"#10",0 },
                        { @"#5",0  },
                        { @"#2",0  },
                        { @"#1",0  },
                        { @"50k",0  },
                        { @"25k",0  },
                        { @"10k",0  },
                        { @"1k",0  },
                        { @"TOTAL",0  }

                    }
                },
                 new NamedCashDenomnation{ Name="AWOLOWO",
                    CashDenomination=new Dictionary<string, double>{
                        { @"#1000",0  },
                        { @"#500",0  },
                        { @"#200",0},
                        { @"#100",0},
                        { @"#50",0},
                        { @"#20",0},
                        { @"#10",0 },
                        { @"#5",0  },
                        { @"#2",0  },
                        { @"#1",0  },
                        { @"50k",0  },
                        { @"25k",0  },
                        { @"10k",0  },
                        { @"1k",0  },
                        { @"TOTAL",0  }


                    }
                 },
                new NamedCashDenomnation{ Name="AWOLOWO",
                    CashDenomination=new Dictionary<string, double>{
                        { @"#1000",0  },
                        { @"#500",0  },
                        { @"#200",0},
                        { @"#100",0},
                        { @"#50",0},
                        { @"#20",0},
                        { @"#10",0 },
                        { @"#5",0  },
                        { @"#2",0  },
                        { @"#1",0  },
                        { @"50k",0  },
                        { @"25k",0  },
                        { @"10k",0  },
                        { @"1k",0  },
                        { @"TOTAL",0  }
                    }
                }
            },null,spreedSheet);
            package.Save();
            return file;
        }

        private FileInfo GenerateConsolidatedReport(){
            var fileName = @"consolidated-report.xlsx";
            FileInfo file = new FileInfo(Path.Combine(_hostingEnvironment.WebRootPath, fileName));
            var package = ExcelReportHelper.CreateExcelPackage(file);
            var spreedSheet = ExcelReportHelper.CreateWorkSheet(package, "Report");
            var startRow = 2;
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
            package.Save();
            return file;
        }
        
        private FileInfo GenerateOtherReport() {
            var fileName = @"others-report.xlsx";
            FileInfo file = new FileInfo(Path.Combine(_hostingEnvironment.WebRootPath, fileName));
            var package = ExcelReportHelper.CreateExcelPackage(file);
            var spreedSheet = ExcelReportHelper.CreateWorkSheet(package, "Report");
            var startRow = 2;
            startRow = ExcelReportHelper.AddSealDecription(startRow, new List<SealDescription>() { new SealDescription { AuditTrail="First of july\n brake shore",
            Seal="1006",SortedValue="FEWFEW",DeclearedValue="DWEDF",Client="FCDF",ATM="DEWD"} }, spreedSheet);

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddDominationProcess(startRow, new Dictionary<string, Domination> {

                {ExcelReportHelper.NairaSymbol+"1000", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"1000",Box=4,Counterfeit=0,Shortages=10,Mixup=0,Overages=0,Total=10} },
                {ExcelReportHelper.NairaSymbol+"500", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"500",Box=1,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {ExcelReportHelper.NairaSymbol+"200", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"200",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {ExcelReportHelper.NairaSymbol+"100", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"100",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {ExcelReportHelper.NairaSymbol+"50", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"50",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {ExcelReportHelper.NairaSymbol+"20", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"20",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {ExcelReportHelper.NairaSymbol+"10", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"10",Box=0,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {ExcelReportHelper.NairaSymbol+"5", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"5",Box=5,Counterfeit=0,Shortages=0,Mixup=0,Overages=0,Total=0} },
                {"TOTAL", new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"5",Box=5,Counterfeit=0,Shortages=10,Mixup=0,Overages=0,Total=10} },
                {"VALUE "+ExcelReportHelper.NairaSymbol, new Domination{ Denomination=ExcelReportHelper.NairaSymbol+"5",Box=17990000,Counterfeit=0,Shortages=10000,Mixup=0,Overages=0,Total=10000} },
            }, spreedSheet);

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddProcessingHeader(startRow, DateTime.Now, "Sterling Bank Plc", DateTime.Now, spreedSheet);

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddBriefSummary(startRow, DateTime.Now, DateTime.Now, "Heritage Banking Company Ltd", ExcelReportHelper.NairaSymbol + "200000000", new Dictionary<string, double> { { "1000", 15000 }, { "500", 5000000 } }, spreedSheet);

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddProcessingDetail(startRow, "170619B1346411", DateTime.Now, DateTime.Now, "K. abayomi", "", spreedSheet);

            startRow = ExcelReportHelper.AddSpace(startRow, 2, spreedSheet);
            startRow = ExcelReportHelper.AddSortingSummary(startRow, 20000000, 20000000, 0, 0, 0, 0, spreedSheet);
            package.Save();
            return file;
        }

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
                    file = GenerateBrokerReport();
                    break;
                case "Consolidated":
                    file = GenerateConsolidatedReport();
                    break;
                case "Daily-Transaction":
                    file = GenerateDailyTransactionReport();
                    break;
                default:
                    file = GenerateOtherReport();
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
