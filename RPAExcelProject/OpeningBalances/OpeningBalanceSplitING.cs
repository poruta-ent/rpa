using OfficeOpenXml;

namespace RPAExcelProject
{
    //TODO V2 testy i refaktor
    public static class OpeningBalaceSplitING
    {
        public static void InsertOpeningBalanceSplitING(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            for (int row = 1; row <= srcSheet.Dimension.End.Row; row++)
            {
                if (srcSheet.Cells[row, 2].Value != null
                    && srcSheet.Cells[row, 8].GetNotNullString() == "Rachunek VAT"
                    && CompanyNamesMapping.DailyAsKey.ContainsKey(srcSheet.Cells[row, 2].Value.ToString().ToUpper())
                    //&& Utils.CompareStringDates(srcSheet.Cells[row, 1].Value.ToString(), MasterData.reportDate)
                    )
                {
                    string lineKey = MasterData.GetCfLineKey(accountBank, 
                                                                srcSheet.Cells[row, 9].GetNotNullString(),
                                                                CompanyNamesMapping.DailyAsKey[srcSheet.Cells[row, 2].GetNotNullString()].CFName);
                    if (MasterData.cfReportLines.ContainsKey(lineKey))
                    {
                        destSheet.Cells[MasterData.cfReportLines[lineKey].RowInCfReport, 5].Value = srcSheet.Cells[row, 6].GetNotNullFloat();
                        //TODO V2 Dodać metodę wstawiającą dla wszystkich plików (przekazujemy arkusz, kolumnę i dane)
                    }
                    else
                    {
                        //TODO V3 Info do usera, że nie ma mapowania
                    }
                }
            }

            for (int row = 1; row <= destSheet.Dimension.End.Row; row ++)
            {
                if (destSheet.Cells[row, 2].GetNotNullString().ToUpper() == "PLN SPLIT"
                    && destSheet.Cells[row, 5].GetNotNullFloat() == 0)
                {
                    destSheet.Cells[row, 5].Value = 0;
                }
            }
        }
        public static void InsertOpeningBalanceSplitINGEscrow(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            for (int row = 1; row <= srcSheet.Dimension.End.Row; row++)
            {
                if (srcSheet.Cells[row, 2].Value != null
                    && srcSheet.Cells[row, 8].GetNotNullString() == "Rachunek escrow z VAT"
                    && CompanyNamesMapping.DailyAsKey.ContainsKey(srcSheet.Cells[row, 2].Value.ToString().ToUpper())
                    //&& Utils.CompareStringDates(srcSheet.Cells[row, 1].Value.ToString(), MasterData.reportDate)
                    )
                {
                    //TODO V1: Poniżej sztywniak przed prezentacją, koniecznie zaimplementować poprawnie rachunki ESCROW w cfReportLines
                    destSheet.Cells[39, 5].Value = srcSheet.Cells[row, 6].GetNotNullFloat();
                }
            }

        }
    }
}
