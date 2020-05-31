using OfficeOpenXml;

namespace RPAExcelProject
{
    public static class OpeningBalance
    {
        public static void InsertOpeningBalanceING(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            int colKey = 2;
            //int colDate = 4;
            int colDescription = 6;
            int colValueToInsert = 3;
            int colCurrency = 4;
            int colDestination = 5;

            for (int row = 1; row <= srcSheet.Dimension.End.Row; row++)
            {
                if (srcSheet.Cells[row, colKey].Value != null
                    && CompanyNamesMapping.DailyAsKey.ContainsKey(srcSheet.Cells[row, colKey].Value.ToString().ToUpper())
                    //&& Utils.CompareStringDates(srcSheet.Cells[row, colDate].Value.ToString(), MasterData.reportDate) //Plik jest robiony na datę, nie potrzebujemy tego warunku.
                    && srcSheet.Cells[row, colDescription].Value.ToString().ToUpper().Contains("TRANSFER ODWROTNY")
                    )
                {
                    string lineKey = MasterData.GetCfLineKey(accountBank, srcSheet.Cells[row, colCurrency].Value.ToString(), CompanyNamesMapping.DailyAsKey[srcSheet.Cells[row, colKey].Value.ToString()].CFName);
                    if (MasterData.cfReportLines.ContainsKey(lineKey))
                    {
                        destSheet.Cells[MasterData.cfReportLines[lineKey].RowInCfReport, colDestination].Value = srcSheet.Cells[row, colValueToInsert].Value;
                        //TODO V3 Obsłużyć błędne wartości w parse
                        //TODO V2 Dodać metodę wstawiającą dla wszystkich plików (przekazujemy arkusz, kolumnę i dane)
                    }
                    else
                    {
                        //TODO V3 Info do usera, że nie ma mapowania
                    }
                }
            }

            // KK moja wersja wypełnienie z OpeningBalanceINGPL kolumna E raportu
            //  wykonanie tej wersji trwa dłużej, gdy chodzimy po arkuszu destSheet
            //OpeningBalanceINGDataProcessor.INGProcessData(MasterData.reportDate, destSheet, srcSheet);
        }

        //TODO V1 - Sztywniak
        public static void InsertOpeningBalanceTFI(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            int colValueToInsert = 4;
            int colDestination = 5;

            destSheet.Cells[144, colDestination].Value = srcSheet.Cells[3, colValueToInsert].GetNotNullFloat();
            destSheet.Cells[140, colDestination].Value = srcSheet.Cells[4, colValueToInsert].GetNotNullFloat();
            destSheet.Cells[139, colDestination].Value = srcSheet.Cells[5, colValueToInsert].GetNotNullFloat();
        }

        //TODO V1 - Sztywniak
        public static void InsertOpeningBalanceFizan(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            int colValueToInsert = 6;
            int colDestination = 5;

            destSheet.Cells[157, colDestination].Value = srcSheet.Cells[2, colValueToInsert].GetNotNullFloat();
            destSheet.Cells[156, colDestination].Value = srcSheet.Cells[5, colValueToInsert].GetNotNullFloat();
            destSheet.Cells[155, colDestination].Value = srcSheet.Cells[7, colValueToInsert].GetNotNullFloat();

        }

    }
}
