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
    }
}