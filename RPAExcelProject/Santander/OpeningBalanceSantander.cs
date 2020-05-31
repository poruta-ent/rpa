using System;
using OfficeOpenXml;

namespace RPAExcelProject
{
    class OpeningBalanceSantander
    {
        public static void InsertDataSantander(ExcelWorksheet destSheet, string nameBank)
        {
            ExcelWorksheet srcSheet = ReportFile.GetWorkbookCsv(ReportFile.OpeningBalanceSantander, MasterData.reportDate, nameBank);

            int colAccount = 2;
            int colCurrency = 9;
            //int colDate = 5;
            int colValueToInsert = 6;
            int colDestination = 5;

            //foreach (KeyValuePair<string, CompanyBankAccount> kvp in MasterData.cfReportLines)
            foreach (var kvp in MasterData.cfReportLines)
            {
                if (kvp.Value.Bank.Contains(nameBank))
                {
                    for (int row = srcSheet.Dimension.Start.Row + 1; row <= srcSheet.Dimension.End.Row; row++)
                    {
                        if (srcSheet.Cells[row, colAccount].Value != null)
                        {
                            //var _bookingDate = DateTime.Parse(srcSheet.Cells[row, colDate].Value.ToString());

                            if ("PL"+srcSheet.Cells[row, colAccount].GetNotNullString().Replace(" ", "") == kvp.Value.AccountNumber.Replace(" ", "")
                                && srcSheet.Cells[row, colCurrency].GetNotNullString() == kvp.Value.Currency
                                //&& _bookingDate.ToString("dd.MM.yyyy") == MasterData.previousReportDate.ToString("dd.MM.yyyy")
                                )
                            {
                                double _amount;
                                var _tmpAmount = srcSheet.Cells[row, colValueToInsert].Value;
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
