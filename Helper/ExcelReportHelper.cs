using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
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
    }
}