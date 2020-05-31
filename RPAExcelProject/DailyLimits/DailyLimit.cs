using System;
using OfficeOpenXml;

namespace RPAExcelProject
{
    public static class DailyLimit
    {
        public static void InsertDailyLimitING(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            int colKey = 4;
            int colDate = 1;
            int colValueToInsert = 2;
            int colCurrency = 3;
            int colDestination = 10;

            for (int row = 1; row <= srcSheet.Dimension.End.Row; row++)
            {
                if (srcSheet.Cells[row, colKey].Value != null
                    && CompanyNamesMapping.DailyAsKey.ContainsKey(srcSheet.Cells[row, colKey].Value.ToString().ToUpper())
                    && Utils.CompareStringDates(srcSheet.Cells[row, colDate].GetNotNullString(), MasterData.reportDate))
                {
                    string lineKey = MasterData.GetCfLineKey(accountBank, srcSheet.Cells[row, colCurrency].Value.ToString(),
                        CompanyNamesMapping.DailyAsKey[srcSheet.Cells[row, colKey].Value.ToString().ToUpper()].CFName);
                    if (MasterData.cfReportLines.ContainsKey(lineKey))
                    {
                        destSheet.Cells[MasterData.cfReportLines[lineKey].RowInCfReport, colDestination].Value = srcSheet.Cells[row, colValueToInsert].Value;
                    }
                    else
                    {
                        //TODO V3 Info do usera, że nie ma mapowania
                    }
                }
            }
        }
    }
}
