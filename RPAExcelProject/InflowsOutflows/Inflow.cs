using OfficeOpenXml;

namespace RPAExcelProject
{
    public static class Inflow
    {
        public static void InsertInflowsING(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            int colKey = 12;
            int colGross = 6;
            int colCurrency = 8;
            int colInflowInDest = 7;
            
            for (int row = 1; row <= srcSheet.Dimension.End.Row; row++)
            {
                if (srcSheet.Cells[row, colKey].Value != null)
                {
                    if (CompanyNamesMapping.InflowsAsKey.ContainsKey(srcSheet.Cells[row, colKey].Value.ToString().ToUpper()))
                    {
                        string lineKey = MasterData.GetCfLineKey(accountBank,
                                                                    srcSheet.Cells[row, colCurrency].GetNotNullString(),
                                                                    CompanyNamesMapping.InflowsAsKey[srcSheet.Cells[row, colKey].GetNotNullString().ToUpper()].CFName);

                        if (MasterData.cfReportLines.ContainsKey(lineKey))
                        {
                            float inflowToInsert = srcSheet.Cells[row, colGross].GetNotNullFloat();
                            int rowToInsert = MasterData.cfReportLines[lineKey].RowInCfReport;

                            if (inflowToInsert != 0)
                                destSheet.Cells[rowToInsert, colInflowInDest].Value = inflowToInsert + destSheet.Cells[rowToInsert, colInflowInDest].GetNotNullFloat();
                        }
                    }
                }
            }
        }
    }
}
