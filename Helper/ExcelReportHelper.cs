using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using razorReport.Models;

namespace razorReport.Helper
{
    public class ExcelReportHelper{

        public readonly static string NairaSymbol = ((char)8358).ToString();
        public static ExcelPackage CreateExcelPackage(FileInfo file)
        {
            //If file exist delete before proceeding.
            if (file.Exists)
            {
                file.Delete();
            }
            var package = new ExcelPackage(file);
            return package;
        }

        public static ExcelWorksheet CreateWorkSheet(ExcelPackage package,string workSheetName)
        {           
            var spreedSheet = package.Workbook.Worksheets.Add(workSheetName);
            return spreedSheet;
        }

        public static int AddSpace(int startRow,int numberOfSpace,ExcelWorksheet worksheet)
        {
            worksheet.Cells[startRow, 1, startRow + numberOfSpace, 20].Merge = true;
            return startRow + numberOfSpace+1;
        }

        public static int AddProcessingHeader(int startRow,DateTime processingTime,string bankName,DateTime deposiTime,ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[startRow, 1, startRow, 3].Merge = true;
            excelWorkSheet.Cells[startRow, 4, startRow, 7].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Value="DATE OF PROCESSING";
            excelWorkSheet.Cells[startRow, 4].Value = processingTime.ToString();

            excelWorkSheet.Cells[startRow+1, 1, startRow+1, 3].Merge = true;
            excelWorkSheet.Cells[startRow+1, 4, startRow+1, 7].Merge = true;
            excelWorkSheet.Cells[startRow+1, 1].Value = "BANK NAME";
            excelWorkSheet.Cells[startRow+1, 4].Value = bankName;

            excelWorkSheet.Cells[startRow + 2, 1, startRow + 2, 3].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 4, startRow + 2, 7].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1].Value = "DATE OF DEPOSIT";
            excelWorkSheet.Cells[startRow + 2, 4].Value = deposiTime.ToString();

            return startRow + 3;
        }

        public static int AddProcessingDetail(int startRow,string BWAuditTrail,DateTime processingStartDate,DateTime processingStopDate,string bankRepresentative,string comment,ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[startRow, 1, startRow, 15].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Value = "PROCESSING DETAILS";
            excelWorkSheet.Row(startRow).Height = 40;
            excelWorkSheet.Cells[startRow, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            excelWorkSheet.Cells[startRow, 1].Style.Font.Bold =true ;
            excelWorkSheet.Cells[startRow, 1].Style.Font.Size = 15;


            startRow = startRow + 1;

            excelWorkSheet.Cells[startRow, 1, startRow, 6].Merge = true;
            excelWorkSheet.Cells[startRow, 7, startRow, 15].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Value = "BW Audit Trail";
            excelWorkSheet.Cells[startRow, 7].Value = BWAuditTrail;


            excelWorkSheet.Cells[startRow+1, 1, startRow+1, 6].Merge = true;
            excelWorkSheet.Cells[startRow+1, 7, startRow+1,15].Merge = true;
            excelWorkSheet.Cells[startRow+1, 1].Value = "Processing Date";
            excelWorkSheet.Cells[startRow+1, 7].Value = processingStartDate.ToString()+" - "+processingStopDate.ToString();

            excelWorkSheet.Cells[startRow + 2, 1, startRow + 2, 6].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 7, startRow + 2, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1].Value = "Bank Representative Witness";
            excelWorkSheet.Cells[startRow+2, 7].Value = bankRepresentative;


            excelWorkSheet.Cells[startRow+3, 1, startRow+3, 15].Merge = true;
            excelWorkSheet.Cells[startRow+3, 1].Value = "Comment :  "+comment;


            return startRow + 4;
        }

        public static int AddBriefSummary(int startRow,DateTime reportDate,DateTime depositDate,string depositBank,string declearedValue,Dictionary<string,double> denominations,ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[startRow, 1, startRow, 15].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Value = "BRIEF SUMMARY";
            excelWorkSheet.Row(startRow).Height = 40;
            excelWorkSheet.Cells[startRow, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            excelWorkSheet.Cells[startRow, 1].Style.Font.Bold = true;
            excelWorkSheet.Cells[startRow, 1].Style.Font.Size = 15;
            string nairaSymbol = ((char)8358).ToString();

            startRow = startRow + 1;

            var denominationsString = "";
            foreach(var deno in denominations)
            {
                denominationsString=denominationsString + deno.Key + ": " +nairaSymbol+deno.Value.ToString("#,##0") + " ";
            }
            excelWorkSheet.Cells[startRow, 1, startRow, 7].Merge = true;
            excelWorkSheet.Cells[startRow, 8, startRow, 15].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Value ="Report Date:";
            excelWorkSheet.Cells[startRow, 8].Value ="   "+reportDate.ToString();
            excelWorkSheet.Cells[startRow, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            excelWorkSheet.Cells[startRow+1, 1, startRow+1, 7].Merge = true;
            excelWorkSheet.Cells[startRow + 1, 8, startRow + 1, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 1, 1].Value = "Deposit Date:";
            excelWorkSheet.Cells[startRow + 1, 8].Value ="   "+ depositDate.ToString();
            excelWorkSheet.Cells[startRow+1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            excelWorkSheet.Cells[startRow + 2, 1, startRow + 2, 7].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 8, startRow + 2, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1].Value = "Deposit Bank:";
            excelWorkSheet.Cells[startRow + 2, 8].Value = "  "+ depositBank;
            excelWorkSheet.Cells[startRow + 2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 7].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 8, startRow + 3, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 1].Value = "Decleared Value:";
            excelWorkSheet.Cells[startRow + 3, 8].Value = "  " + declearedValue;
            excelWorkSheet.Cells[startRow + 3, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            excelWorkSheet.Cells[startRow + 4, 1, startRow + 4, 7].Merge = true;
            excelWorkSheet.Cells[startRow + 4, 8, startRow + 4, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 4, 1].Value = "Denominations:";
            excelWorkSheet.Cells[startRow + 4, 8].Value = "  "+denominationsString;
            excelWorkSheet.Cells[startRow + 4, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            return startRow + 5;
        }

        public static int AddSealDecription(int startRow, List<SealDescription> sealDecriptions,ExcelWorksheet spreedSheet)
        {
            spreedSheet.Cells["A" + startRow.ToString()].Value = "Seal";

            spreedSheet.Cells["B" + startRow.ToString()].Value = "Audit Trail";
            spreedSheet.Cells[startRow, 2, startRow, 3].Merge = true;

            spreedSheet.Cells[startRow,4].Value = "Client";
            spreedSheet.Cells[startRow, 4, startRow, 5].Merge = true;

            spreedSheet.Cells[startRow,6].Value = "Declared Value";
            spreedSheet.Cells[startRow, 6, startRow, 7].Merge = true;

            spreedSheet.Cells[startRow,8].Value = "Sorted Value";
            spreedSheet.Cells[startRow, 8, startRow, 9].Merge = true;

            spreedSheet.Cells[startRow,10].Value = "Analysis";
            spreedSheet.Cells[startRow, 10, startRow, 11].Merge = true;

            spreedSheet.Cells["A" + startRow.ToString() + ":Z" + startRow.ToString()].Style.Font.Bold = true;
            startRow = startRow + 1;
            
            foreach (var sealDescription in sealDecriptions)
            {
                spreedSheet.Cells[startRow,1].Value = sealDescription.Seal;
                spreedSheet.Cells[startRow, 1, startRow + 5, 1].Merge = true;
                spreedSheet.Cells["A" + startRow.ToString()].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                spreedSheet.Cells[startRow,2].Value = sealDescription.AuditTrail;
                spreedSheet.Cells[startRow, 2, startRow + 4, 3].Merge = true;
                spreedSheet.Cells[startRow,2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                spreedSheet.Cells[startRow, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Justify;

                spreedSheet.Cells[startRow,4].Value = sealDescription.Client;
                spreedSheet.Cells[startRow, 4, startRow + 4, 5].Merge = true;
                spreedSheet.Cells[startRow,4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                spreedSheet.Cells[startRow,6].Value = sealDescription.DeclearedValue;
                spreedSheet.Cells[startRow, 6, startRow + 4, 7].Merge = true;
                spreedSheet.Cells[startRow,6].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                spreedSheet.Cells[startRow,8].Value = sealDescription.SortedValue;
                spreedSheet.Cells[startRow, 8, startRow + 4, 9].Merge = true;
                spreedSheet.Cells[startRow,8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                spreedSheet.Cells[startRow,10].Value = sealDescription.ATM; 
                startRow = startRow + 6;
            }
         
            return startRow;
        }

        public static int AddDailyConsolidatedReport(int startRow, Dictionary<string, CurrencyType> currencies, ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Merge = true;
            excelWorkSheet.Cells[startRow + 1, 1, startRow + 1, 20].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1, startRow + 2, 20].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 20].Merge = true;

            excelWorkSheet.Cells[startRow, 1].Value = "BANKERS WEARHOUSE PLC";
            excelWorkSheet.Cells[startRow, 1].Style.Font.Bold = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);


            excelWorkSheet.Cells[startRow + 1, 1].Value = "CONSOLIDATED REPORT";
            excelWorkSheet.Cells[startRow + 1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow + 1, 1].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            excelWorkSheet.Cells[startRow + 2, 1].Value = "STANBIC DAILY CASH OPERATION REPORT (NGN)";
            excelWorkSheet.Cells[startRow + 2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow + 2, 1].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            excelWorkSheet.Cells[startRow + 3, 1].Value = DateTime.Now.ToString();
            excelWorkSheet.Cells[startRow + 3, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow + 3, 1].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            excelWorkSheet.Cells[startRow, 1, startRow + 3, 20].Style.Font.Color.SetColor(Color.White);
            excelWorkSheet.Cells[startRow + 4, 1, startRow + 4, 20].Merge = true;

            startRow = startRow + 5;
            int column = 5;



            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Font.Color.SetColor(Color.White);

            excelWorkSheet.Cells[startRow + 1, 1, startRow + 1, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 1, 1].Value = "MINT";
            excelWorkSheet.Cells[startRow + 2, 1, startRow + 2, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1].Value = "ATM";
            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 1].Value = "CAC";
            excelWorkSheet.Cells[startRow + 4, 1, startRow + 4, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 4, 1].Value = "CAD";
            excelWorkSheet.Cells[startRow + 5, 1, startRow + 5, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 5, 1].Value = "AE";
            excelWorkSheet.Cells[startRow + 6, 1, startRow + 6, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 6, 1].Value = DateTime.Now.ToString();
            excelWorkSheet.Cells[startRow + 7, 1, startRow + 7, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 7, 1].Value = "CITI BANK";
            excelWorkSheet.Cells[startRow + 8, 1, startRow + 8, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 8, 1].Value = "FIDELITY BANK";
            excelWorkSheet.Cells[startRow + 9, 1, startRow + 9, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 9, 1].Value = "FCMB";
            excelWorkSheet.Cells[startRow + 10, 1, startRow + 10, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 10, 1].Value = "UBA";
            excelWorkSheet.Cells[startRow + 11, 1, startRow + 11, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 11, 1].Value = "DIAMOND BANK";
            excelWorkSheet.Cells[startRow + 12, 1, startRow + 12, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 12, 1].Value = "KANO SWAP";
            excelWorkSheet.Cells[startRow + 13, 1, startRow + 13, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 13, 1].Value = "IBADAN SWAP";
            excelWorkSheet.Cells[startRow + 14, 1, startRow + 14, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 14, 1].Value = "AT COB";

            foreach (var currency in currencies)
            {
                excelWorkSheet.Cells[startRow, column].Value = currency.Key;
                excelWorkSheet.Cells[startRow + 1, column].Value = currency.Value.Mint.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 1, column].AutoFitColumns();
                excelWorkSheet.Cells[startRow + 2, column].Value = currency.Value.ATM.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 3, column].Value = currency.Value.CAC.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 4, column].Value = currency.Value.CAD.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 5, column].Value = currency.Value.AE.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 6, column].Value = currency.Value.Today.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 7, column].Value = currency.Value.CITI.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 8, column].Value = currency.Value.Fidelity.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 9, column].Value = currency.Value.FCMB.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 10, column].Value = currency.Value.UBA.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 11, column].Value = currency.Value.DIAMOND.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 12, column].Value = currency.Value.KANOSWAP.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 13, column].Value = currency.Value.IBADANSWAP.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 14, column].Value = currency.Value.ATCOB.ToString("#,##0");

                column = column + 1;

            }

            return startRow + 15;
        }

        public static int AddDominationProcess(int startRow, Dictionary<string, Domination> dominations, ExcelWorksheet workSheet)
        {
            workSheet.Cells[startRow, 1, startRow, 3].Merge = true;
            workSheet.Cells[startRow + 1, 1, startRow + 1, 3].Merge = true;
            workSheet.Cells[startRow + 2, 1, startRow + 2, 15].Merge = true;

            workSheet.Cells[startRow + 2, 1].Style.Font.Bold = true;
            workSheet.Cells[startRow + 2, 1].Style.Font.Size = 15;

            workSheet.Cells[startRow + 3, 1, startRow + 3, 3].Merge = true;
            workSheet.Cells[startRow + 4, 1, startRow + 4, 3].Merge = true;
            workSheet.Cells[startRow + 5, 1, startRow + 5, 3].Merge = true;
            workSheet.Cells[startRow + 6, 1, startRow + 6, 3].Merge = true;
            workSheet.Cells[startRow + 7, 1, startRow + 7, 3].Merge = true;

            workSheet.Cells[startRow, 1].Value = "DENOMINATION PROCESSED";
            workSheet.Cells[startRow + 1, 1].Value = "BOXES PROCESSED";
            workSheet.Cells[startRow + 2, 1].Value = "DESCRIPANCIES";
            workSheet.Cells[startRow + 3, 1].Value = "SHORTAGES (Pcs.)";
            workSheet.Cells[startRow + 4, 1].Value = "MIX UPS (Pcs.)";
            workSheet.Cells[startRow + 5, 1].Value = "OVERAGES (Pcs)";
            workSheet.Cells[startRow + 6, 1].Value = "COUNTERFEITS (Pcs)";
            workSheet.Cells[startRow + 7, 1].Value = "TOTAL";

            int count = 4;


            foreach (var domination in dominations)
            {
                workSheet.Cells[startRow, count].Value = domination.Key;
                workSheet.Cells[startRow + 1, count].Value = domination.Value.Box == 0 ? "Nil" : domination.Value.Box.ToString("#,##0");
                workSheet.Cells[startRow + 3, count].Value = domination.Value.Shortages == 0 ? "" : domination.Value.Shortages.ToString("#,##0");
                workSheet.Cells[startRow + 4, count].Value = domination.Value.Mixup == 0 ? "" : domination.Value.Mixup.ToString("#,##0");
                workSheet.Cells[startRow + 5, count].Value = domination.Value.Overages == 0 ? "" : domination.Value.Overages.ToString("#,##0");
                workSheet.Cells[startRow + 6, count].Value = domination.Value.Counterfeit == 0 ? "" : domination.Value.Counterfeit.ToString("#,##0");
                workSheet.Cells[startRow + 7, count].Value = domination.Value.Total == 0 ? "" : domination.Value.Total.ToString("#,##0");
                count = count + 1;
            }

            return startRow + 8;
        }

        public static int AddSortingSummary(int startRow,double declearedValue,double postSortingValue,double shortage,double overages,double counterfeits,double Mixup,ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[startRow, 1, startRow, 15].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Value = "SORTING SUMMARY";
            excelWorkSheet.Row(startRow).Height = 40;
            excelWorkSheet.Cells[startRow, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            excelWorkSheet.Cells[startRow, 1].Style.Font.Bold = true;
            excelWorkSheet.Cells[startRow, 1].Style.Font.Size = 15;
            startRow = startRow + 1;


            excelWorkSheet.Cells[startRow, 1, startRow, 8].Merge = true;
            excelWorkSheet.Cells[startRow, 9, startRow, 15].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Value = "Decleared Value:";
            excelWorkSheet.Cells[startRow, 9].Value =NairaSymbol+ declearedValue.ToString("#,##0");



            excelWorkSheet.Cells[startRow+1, 1, startRow+1, 8].Merge = true;
            excelWorkSheet.Cells[startRow+1, 9, startRow+1, 15].Merge = true;
            excelWorkSheet.Cells[startRow+1, 1].Value = "Post Sorting Value:";
            excelWorkSheet.Cells[startRow+1, 9].Value =NairaSymbol+ postSortingValue.ToString("#,##0");

            excelWorkSheet.Cells[startRow + 2, 1, startRow + 2, 8].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 9, startRow + 2, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1].Value = "Shortages:";
            excelWorkSheet.Cells[startRow + 2, 9].Value =NairaSymbol+ shortage.ToString("#,##0");


            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 8].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 9, startRow + 3, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 1].Value = "Overages:";
            excelWorkSheet.Cells[startRow + 3, 9].Value = NairaSymbol+ overages.ToString("#,##0");


            excelWorkSheet.Cells[startRow + 4, 1, startRow + 4, 8].Merge = true;
            excelWorkSheet.Cells[startRow + 4, 9, startRow + 4, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 4, 1].Value = "Counterfeits:";
            excelWorkSheet.Cells[startRow + 4, 9].Value =NairaSymbol+ counterfeits.ToString("#,##0");

            excelWorkSheet.Cells[startRow + 5, 1, startRow + 5, 8].Merge = true;
            excelWorkSheet.Cells[startRow + 5, 9, startRow + 5, 15].Merge = true;
            excelWorkSheet.Cells[startRow + 5, 1].Value = "Mix-ups:";
            excelWorkSheet.Cells[startRow + 5, 9].Value =NairaSymbol +Mixup.ToString("#,##0");

            return startRow + 7;
        }

        public static int AddDailyTransaction(int startRow,Dictionary<string,double> bankOpenning,List<NamedCashDenomnation> inflowEvacuation,List<NamedCashDenomnation> unNamedCashDenomnations, ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Merge = true;
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Font.Color.SetColor(Color.Black);
            excelWorkSheet.Cells[startRow, 1].Value = "DAILY TRANSACTIONS LOCAL CURRENCY (NGN)";

            excelWorkSheet.Cells[startRow + 1, 1, startRow+1, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1, startRow+2, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 4].Merge = true;

            excelWorkSheet.Cells[startRow + 1, 1].Value = "";
            excelWorkSheet.Cells[startRow + 2, 1].Value = "BANK OPENNING BALANCE";
            excelWorkSheet.Cells[startRow + 3, 1].Value = "INFLOW/EVACUATION";
            int column = 5;
        
            foreach(var bankOp in bankOpenning)
            {
                excelWorkSheet.Cells[startRow + 1, column].Value = bankOp.Key;
                excelWorkSheet.Cells[startRow + 2, column].Value = bankOp.Value.ToString("#,##0");
              
                column =column+ 1;
            }
            excelWorkSheet.Cells[startRow+2, 1, startRow+2, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow+2, 1, startRow+2, 20].Style.Fill.BackgroundColor.SetColor(Color.Black);
            excelWorkSheet.Cells[startRow+2, 1, startRow+2, 20].Style.Font.Color.SetColor(Color.White);
            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 20].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            startRow= startRow + 4;
            startRow = AddSubSection(startRow, inflowEvacuation, excelWorkSheet);

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
            excelWorkSheet.Cells[startRow, 5, startRow, 20].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType=ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;        
            excelWorkSheet.Cells[startRow, 1].                                                                                                                                                                                                                                                                                                                                                                                                                                            Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1].Value = "Mutilated Evacuation";


            // startRow = AddSubSection(startRow, unNamedCashDenomnations, excelWorkSheet);
            startRow = startRow + 2;

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1].Value = "Evacuation in BW Wrapper";

            startRow = startRow + 2;

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1].Value = "Evacuation in PAPER NOTE";

            startRow = startRow + 2;

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1].Value = "TO BE RETURNED TO BRANCH";

            startRow = startRow + 2;

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1].Value = "RETURNED TO VAULT";

            startRow = startRow + 2;

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1].Value = "CASH SWAP";

            startRow = startRow + 2;

            excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
            excelWorkSheet.Cells[startRow+1, 1, startRow+1, 4].Merge = true;
            excelWorkSheet.Cells[startRow, 1,startRow+1,4].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1,startRow+1,4].Style.Fill.BackgroundColor.SetColor(Color.Gray);
            excelWorkSheet.Cells[startRow, 1].Value = "OTHERS:";
            excelWorkSheet.Cells[startRow+1, 1].Value = "TRANSFER FROM ILUPEJU";

            startRow = startRow + 3;
            return startRow;
        }

        private static int AddSubSection(int startRow, List<NamedCashDenomnation> subsectionList,ExcelWorksheet excelWorkSheet)
        {
            foreach (var ife in subsectionList)
            {
                excelWorkSheet.Cells[startRow, 1, startRow, 4].Merge = true;
                excelWorkSheet.Cells[startRow, 1].Value = ife.Name;
                excelWorkSheet.Cells[startRow, 5].Value = ife.CashDenomination["#1000"];
                excelWorkSheet.Cells[startRow, 6].Value = ife.CashDenomination["#500"];
                excelWorkSheet.Cells[startRow, 7].Value = ife.CashDenomination["#200"];
                excelWorkSheet.Cells[startRow, 8].Value = ife.CashDenomination["#100"];
                excelWorkSheet.Cells[startRow, 9].Value = ife.CashDenomination["#50"];
                excelWorkSheet.Cells[startRow, 10].Value = ife.CashDenomination["#20"];
                excelWorkSheet.Cells[startRow, 11].Value = ife.CashDenomination["#10"];
                excelWorkSheet.Cells[startRow, 12].Value = ife.CashDenomination["#5"];
                excelWorkSheet.Cells[startRow, 13].Value = ife.CashDenomination["#2"];
                excelWorkSheet.Cells[startRow, 14].Value = ife.CashDenomination["#1"];
                excelWorkSheet.Cells[startRow, 15].Value = ife.CashDenomination["50k"];
                excelWorkSheet.Cells[startRow, 16].Value = ife.CashDenomination["25k"];
                excelWorkSheet.Cells[startRow, 17].Value = ife.CashDenomination["10k"];
                excelWorkSheet.Cells[startRow, 18].Value = ife.CashDenomination["1k"];
                excelWorkSheet.Cells[startRow, 19].Value = ife.CashDenomination["TOTAL"];
                startRow = startRow + 1;
            }
            return startRow;
        }

        public static int AddBankBrokerReport(int startRow,string bankHeading, Dictionary<string, CurrencyType> currencies, ExcelWorksheet excelWorkSheet)
        {
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Merge = true;

            excelWorkSheet.Cells[startRow, 1].Value = bankHeading;
            excelWorkSheet.Cells[startRow, 1].Style.Font.Bold = true;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255));
            excelWorkSheet.Row(startRow).Height = 40;
            excelWorkSheet.Cells[startRow, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            startRow += 1;
            int column = 5;



            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Fill.PatternType = ExcelFillStyle.Solid;
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
            excelWorkSheet.Cells[startRow, 1, startRow, 20].Style.Font.Color.SetColor(Color.White);

            excelWorkSheet.Cells[startRow + 1, 1, startRow + 1, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 1, 1].Value = "MINT";
            excelWorkSheet.Cells[startRow + 2, 1, startRow + 2, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 2, 1].Value = "ATM";
            excelWorkSheet.Cells[startRow + 3, 1, startRow + 3, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 3, 1].Value = "CAC";
            excelWorkSheet.Cells[startRow + 4, 1, startRow + 4, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 4, 1].Value = "CAD";
            excelWorkSheet.Cells[startRow + 5, 1, startRow + 5, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 5, 1].Value = "AE";
            excelWorkSheet.Cells[startRow + 6, 1, startRow + 6, 4].Merge = true;
            excelWorkSheet.Cells[startRow + 6, 1].Value = "At COB";

            foreach (var currency in currencies)
            {
                excelWorkSheet.Cells[startRow, column].Value = currency.Key;
                excelWorkSheet.Cells[startRow + 1, column].Value = currency.Value == null ? "" : currency.Value.Mint.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 1, column].AutoFitColumns();
                excelWorkSheet.Cells[startRow + 2, column].Value = currency.Value == null ? "" : currency.Value.ATM.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 3, column].Value = currency.Value == null ? "" : currency.Value.CAC.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 4, column].Value = currency.Value == null ? "" : currency.Value.CAD.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 5, column].Value = currency.Value == null ? "" : currency.Value.AE.ToString("#,##0");
                excelWorkSheet.Cells[startRow + 6, column].Value = currency.Value == null ? "" : currency.Value.ATCOB.ToString("#,##0");

                column = column + 1;

            }

            return startRow + 7;
        }

    }
}