using System;
using OfficeOpenXml;

namespace RPAExcelProject
{
    public static class Outflow
    {
        public static void InsertOutflows(ExcelWorksheet destSheet, ExcelWorksheet srcSheet, string accountBank)
        {
            int colKey = 5;
            int colTitle = 2;
            int colGross = 6;
            int colVAT = 7;
            int colCurrency = 8;
            int colSplit = 13;
            int colOutflowInDest = 6;
            int colOpeningBalanceInDest = 5;

            for (int row = 1; row <= srcSheet.Dimension.End.Row; row++)
            {
                if (srcSheet.Cells[row, colKey].Value != null)
                {
                    if (CompanyNamesMapping.K2AsKey.ContainsKey(srcSheet.Cells[row, colKey].Value.ToString().ToUpper()))
                    {
                        string lineKey = MasterData.GetCfLineKey(accountBank, 
                                                                    srcSheet.Cells[row, colCurrency].GetNotNullString(),
                                                                    CompanyNamesMapping.K2AsKey[srcSheet.Cells[row, colKey].GetNotNullString().ToUpper()].CFName);

                        if (MasterData.cfReportLines.ContainsKey(lineKey))
                        {

                            float nonSplitValueToInsert;
                            float splitValueToInsert;
                            float splitAccumulated;
                            int rowToInsert = MasterData.cfReportLines[lineKey].RowInCfReport;
                            int rowToInsertSplit = rowToInsert + 1;
                            while (destSheet.Cells[rowToInsertSplit, colTitle].GetNotNullString().ToUpper() != "PLN SPLIT") { rowToInsertSplit++; }
                            
                            if (IsOutflowSplit(srcSheet.Cells[row, colSplit].GetNotNullString(),
                                                srcSheet.Cells[row, colGross].GetNotNullFloat(),
                                                srcSheet.Cells[row, colVAT].GetNotNullFloat(),
                                                srcSheet.Cells[row, colTitle].GetNotNullString()))
                            {

                                splitAccumulated = -1 * destSheet.Cells[rowToInsertSplit, colOutflowInDest].GetNotNullFloat();

                                splitValueToInsert = CalculateSplitOutflow(srcSheet.Cells[row, colVAT].GetNotNullFloat(),
                                                                    splitAccumulated,    
                                                                    destSheet.Cells[rowToInsertSplit, colOpeningBalanceInDest].GetNotNullFloat());

                                nonSplitValueToInsert = srcSheet.Cells[row, colGross].GetNotNullFloat() - splitValueToInsert;
                            }
                            else
                            {
                                splitAccumulated = 0f;
                                splitValueToInsert = 0f;
                                nonSplitValueToInsert = srcSheet.Cells[row, colGross].GetNotNullFloat();
                            }

                            if (splitValueToInsert!= 0)
                                destSheet.Cells[rowToInsertSplit, colOutflowInDest].Value = -1 * splitValueToInsert 
                                                                        + destSheet.Cells[rowToInsertSplit, colOutflowInDest].GetNotNullFloat();

                            if (nonSplitValueToInsert != 0)
                                destSheet.Cells[rowToInsert, colOutflowInDest].Value = -1 * nonSplitValueToInsert 
                                                                        + destSheet.Cells[rowToInsert, colOutflowInDest].GetNotNullFloat();

                            /*if (ReportFile.SplitTest && rowToInsert==19) LogsDisplay.SplitLogsInsert(rowToInsert,
                                                                                    srcSheet.Cells[row, colTitle].GetNotNullString(),
                                                                                    srcSheet.Cells[row, colSplit].GetNotNullString(),
                                                                                    srcSheet.Cells[row, colGross].GetNotNullFloat(),
                                                                                    srcSheet.Cells[row, colVAT].GetNotNullFloat(),
                                                                                    destSheet.Cells[rowToInsertSplit, colOpeningBalanceInDest].GetNotNullFloat(),
                                                                                    splitAccumulated,
                                                                                    splitValueToInsert,
                                                                                    nonSplitValueToInsert
                                                                                ); */
                        }

                        //TODO V2 Obsłużyć inne waluty niż EUR i PLN (zamiana na PLN)

                        else
                        { 
                            //TODO V3 Info do usera, że nie ma mapowania
                        }
                    }
                }
            }
        }


        public static bool IsOutflowSplit(string splitColumnValue, float grossAmount, float VAT, string transactionDescription)
        {
            return splitColumnValue.ToUpper() == "TRUE" || grossAmount == VAT || transactionDescription.Contains("VAT");
        }

        public static float CalculateSplitOutflow(float vat, float splitOutflowAccumulated, float openingBalanceSplit)
        {
            if (openingBalanceSplit < vat + splitOutflowAccumulated)
            {
                return openingBalanceSplit - splitOutflowAccumulated;
            }
            else
            {
                return vat;
            }
        }
    }
}
