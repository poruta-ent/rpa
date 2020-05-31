using System;
using OfficeOpenXml;

namespace RPAExcelProject
{
    class OpeningBalanceSEB
    {
        public static void InsertDataSEB(ExcelWorksheet destSheet, string nameBank)
        {
            ExcelWorksheet srcSEBSheet = ReportFile.GetWorkbookCsv(ReportFile.OpeningBalanceSEB, MasterData.reportDate, nameBank);

            // blokuję, bo wykorzystuję do testów
            KK_TestDict.DisplayMappingsInSheet(MasterData.reportWorkbook);

            int colAccount = 2;
            int colCurrency = 3;
            int colDate = 5;
            int colValueToInsert = 9;
            int colDestination = 5;

            //foreach (KeyValuePair<string, CompanyBankAccount> kvp in MasterData.cfReportLines)
            foreach (var kvp in MasterData.cfReportLines)
            {
                if (kvp.Value.Bank.Contains(nameBank))
                {
                    for (int row = srcSEBSheet.Dimension.Start.Row+1; row<= srcSEBSheet.Dimension.End.Row; row++)
                    {
                        if (srcSEBSheet.Cells[row, colAccount].Value != null)
                        {
                            var _bookingDate = DateTime.Parse(srcSEBSheet.Cells[row, colDate].Value.ToString());

                            if (srcSEBSheet.Cells[row, colAccount].GetNotNullString().Replace(" ", "") == kvp.Value.AccountNumber.Replace(" ", "")
                                && srcSEBSheet.Cells[row, colCurrency].GetNotNullString() == kvp.Value.Currency
                                && _bookingDate.ToString("dd.MM.yyyy") == MasterData.previousReportDate.ToString("dd.MM.yyyy")
                                )
                            {
                                double _amount;
                                var _tmpAmount = srcSEBSheet.Cells[row, colValueToInsert].Value;
                                if (_tmpAmount != null)
                                {
                                    Double.TryParse(_tmpAmount.ToString(), out _amount);
                                    destSheet.Cells[kvp.Value.RowInCfReport, colDestination].Value = _amount;
                                }
                                else
                                    destSheet.Cells[kvp.Value.RowInCfReport, colDestination].Value = 0.0;

                                break;
                            }
                        }
                    }
                }
            }
        }
    }
}
