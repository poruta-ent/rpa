using System.Collections.Generic;

namespace RPAExcelProject
{
    public static class CompanyNamesMapping
    {
        public static Dictionary<string, CompanyName> CFAsKey;
        public static Dictionary<string, CompanyName> K2AsKey;
        public static Dictionary<string, CompanyName> DailyAsKey;
        public static Dictionary<string, CompanyName> NettingAsKey;
        public static Dictionary<string, CompanyName> InflowsAsKey;

        public static Dictionary<string, CompanyName> GetMappings(int keyColumn)
        {
            var namesMapping = new Dictionary<string, CompanyName>();

            using (var mappingWorkbook = ReportFile.GetWorkbook(ReportFile.CompanyMapping))
            {
                var mappingSheet = mappingWorkbook.Workbook.Worksheets[0];

                int row = 2;
                int colKey = keyColumn;
                int colCF = 1;
                int colKorab2 = 2;
                int colDailyFlow = 3;
                int colNetting = 4;
                int colInflows = 5;

                while (mappingSheet.Cells[row, colKey].Value != null)
                {
                    if (!namesMapping.ContainsKey(mappingSheet.Cells[row, colKey].GetNotNullString().ToUpper()))
                    {
                        namesMapping.Add(mappingSheet.Cells[row, colKey].GetNotNullString().ToUpper(),
                                            new CompanyName
                                            {
                                                CFName = mappingSheet.Cells[row, colCF].GetNotNullString().ToUpper(),
                                                Korab2Name = mappingSheet.Cells[row, colKorab2].GetNotNullString().ToUpper(),
                                                DailyFlowName = mappingSheet.Cells[row, colDailyFlow].GetNotNullString().ToUpper(),
                                                NettingName = mappingSheet.Cells[row, colNetting].GetNotNullString().ToUpper(),
                                                InflowsName = mappingSheet.Cells[row, colInflows].GetNotNullString().ToUpper()
                                            });
                    }
                    row++;
                }
            }
            return namesMapping;
        }

        public static void GenerateMappings()
        {
            CFAsKey = GetMappings(1);
            K2AsKey = GetMappings(2);
            DailyAsKey = GetMappings(3);
            NettingAsKey = GetMappings(4);
            InflowsAsKey = GetMappings(5);
        }
    }
}
