using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace RPAExcelProject
{
    public static class LogsDisplay
    {
        public static void SplitLogsInsert(int rowInCF, string desc, string isSplit, float gross, float vat, float obSplit, float accSplit, float splitVal, float splitOver)
        {
            var sh = ReportFile.logWorkbook.Workbook.Worksheets["split"]; 
                
                //ReportFile.splitLogs;
            int row = 2;
            while (sh.Cells[row,1].Value!=null) { row++; }
            sh.Cells[row, 1].Value = rowInCF;
            sh.Cells[row, 2].Value = desc;

            if (desc.ToUpper().Contains("VAT"))
            {
                sh.Cells[row, 3].Value = 1;
            }
            else
            {
                sh.Cells[row, 3].Value = 0;
            }
            

            sh.Cells[row, 4].Value = isSplit;
            sh.Cells[row, 5].Value = gross;
            sh.Cells[row, 6].Value = vat;
            sh.Cells[row, 7].Value = gross - vat;

            if (gross == vat)
            {
                sh.Cells[row, 8].Value = 1;
            }
            else
            {
                sh.Cells[row, 8].Value = 0;
            }

            sh.Cells[row, 9].Value = obSplit;
            sh.Cells[row, 10].Value = accSplit;
            sh.Cells[row, 11].Value = splitVal;
            sh.Cells[row, 12].Value = splitOver;

            ReportFile.logWorkbook.Save();
        }

        /// <summary>
        /// Metoda pomocnicza tworzy arkusz z danymi ze słownika z mapowaniem.
        /// </summary>
        /// <returns></returns>
        public static void CfReportLinesLogsInsert()
        {

            var dictSheet = ReportFile.logWorkbook.Workbook.Worksheets.Where(x => x.Name == "Dict").FirstOrDefault();
            if (dictSheet != null) ReportFile.logWorkbook.Workbook.Worksheets.Delete("Dict");
            dictSheet = ReportFile.logWorkbook.Workbook.Worksheets.Add("Dict");

            dictSheet.Cells[1, 1].Value = "Key";
            dictSheet.Cells[1, 2].Value = "Acc desc";
            dictSheet.Cells[1, 3].Value = "Row in cf rerport";
            dictSheet.Cells[1, 4].Value = "Name in K2";
            dictSheet.Cells[1, 5].Value = "Currency";
            dictSheet.Cells[1, 6].Value = "Bank";
            dictSheet.Cells[1, 7].Value = "Number";

            int row = 2;
            foreach (KeyValuePair<string, CompanyBankAccount> kvp in MasterData.cfReportLines)
            {
                dictSheet.Cells[row, 1].Value = kvp.Key;
                dictSheet.Cells[row, 2].Value = kvp.Value.AccountDesc;
                dictSheet.Cells[row, 3].Value = kvp.Value.RowInCfReport;
                dictSheet.Cells[row, 4].Value = kvp.Value.CompanyNaneInKorab2;
                dictSheet.Cells[row, 5].Value = kvp.Value.Currency;
                dictSheet.Cells[row, 6].Value = kvp.Value.Bank;
                dictSheet.Cells[row, 7].Value = kvp.Value.AccountNumber;
                row++;
            }
            ReportFile.logWorkbook.Save();
        }

    }
}
