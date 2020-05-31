using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace RPAExcelProject
{
    class KK_TestDict
    {
        /// <summary>
        /// Metoda pomocnicza tworzy arkusz z danymi ze słownika z mapowaniem.
        /// </summary>
        /// <returns></returns>

        public static void DisplayMappingsInSheet(ExcelPackage workbook)
        {
            var dictSheet = workbook.Workbook.Worksheets.Where(x => x.Name == "Dict").FirstOrDefault();

            if (dictSheet != null) workbook.Workbook.Worksheets.Delete("Dict");

            dictSheet = workbook.Workbook.Worksheets.Add("Dict");

            dictSheet.Cells[1, 1].Value = "Key";
            dictSheet.Cells[1, 2].Value = "Acc desc";
            dictSheet.Cells[1, 3].Value = "Row in cf rerport";
            dictSheet.Cells[1, 4].Value = "Name in K2";
            dictSheet.Cells[1, 5].Value = "Currency";
            dictSheet.Cells[1, 6].Value = "Bank";
            dictSheet.Cells[1, 7].Value = "Number";
            dictSheet.Cells[1, 8].Value = "Company";

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
                dictSheet.Cells[row, 8].Value = kvp.Value.CompanyNameInBank;
                //Console.WriteLine($"Company: {kvp.Key}\t\t\t\t\t\t\t details:\t\tRow={kvp.Value.RowInCfReport}\t\tK2Name={kvp.Value.CompanyNaneInKorab2}\t\tBank={kvp.Value.Bank}\t\tCurr={kvp.Value.Currency}\t\tType={kvp.Value.AccountType}\t\tNo={kvp.Value.AccountNumber}\t\t");
                row++;
            }

            workbook.Save();
        }
    }
}
