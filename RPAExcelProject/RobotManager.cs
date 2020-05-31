using System;
using System.Diagnostics;
using OfficeOpenXml;

namespace RPAExcelProject
{
    public static class RobotManager
    {
        //TODO V2 Konieczny refaktor
        public static void RunCfRobot()
        { 
            bool processStageStatus;
            string processMessage = string.Empty;

            var robotWatch = Stopwatch.StartNew();
            Console.WriteLine($"Starting process for CashFlow report\n");

            // 2019-08-05 KK trzeba tu przenieść MasterData.reportDate, bo jest wykorzystywany jako część nazwy pliku .csv
            // 2019-08-07 sprawdziłem, że nie ma znaczenia wielkość liter przy rozszerzeniu pliku .CSV czy .csv
            MasterData.reportDate = "09.08.2019";
            
            MasterData.previousReportDate = Utils.PreviousDate(DateTime.Parse(MasterData.reportDate));

            //KK pozyskanie włściwych nazw plików .csv
            Utils.FilesFromCsv();

            Console.WriteLine($"Step 1 - Checking the required files ....");
            var processWatch = Stopwatch.StartNew();
            processStageStatus = ReportFile.CheckIfCFFilesReady(out processMessage);
            if (processStageStatus)
            {
                MasterData.reportWorkbook = ReportFile.GetWorkbook(ReportFile.CFReport);
            }
            processWatch.Stop();
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");
            if (!processStageStatus) return;

            Console.WriteLine($"Step 2 - Configuring report worksheet....");
            processWatch = Stopwatch.StartNew();
            processStageStatus = CanCreateReportSheet();
            if (processStageStatus)
            {
                ReportFile.GenerateLogSheets();
                MasterData.reportSheet = MasterData.CreateFromTemplate();
                // KK dodałem wyświetlenie, jaki arkusz został utworzony
                processMessage = $"Success! Report worksheet {MasterData.reportDate} created.";
            }
            else
            {
                processMessage = "Aborted by user!";
            }
            processWatch.Stop();
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");
            if (!processStageStatus) return;

            Console.WriteLine($"Step 3 - Initializing mappings....");
            processWatch = Stopwatch.StartNew();
            MasterData.InitializeMasterdata();
            processMessage = "Success! Mappings ready.";
            processWatch.Stop();
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");

            Console.WriteLine($"Step 4 - Reading Opening balances ING PL....");
            processWatch = Stopwatch.StartNew();
            var obSheet = ReportFile.GetWorkbook(ReportFile.OpeningBalanceINGPL).Workbook.Worksheets[0];
            OpeningBalance.InsertOpeningBalanceING(MasterData.reportSheet, obSheet, "ING CP");
            obSheet = ReportFile.GetWorkbook(ReportFile.OpeningBalanceTFI).Workbook.Worksheets[0];
            OpeningBalance.InsertOpeningBalanceTFI(MasterData.reportSheet, obSheet, "ING CP");
            obSheet = ReportFile.GetWorkbook(ReportFile.OpeningBalanceFizan).Workbook.Worksheets[0];
            OpeningBalance.InsertOpeningBalanceFizan(MasterData.reportSheet, obSheet, "ING CP");
            processWatch.Stop();
            processMessage = "Success! Opening balances loaded.";
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");

            Console.WriteLine($"Step 5 - Reading Split Payment....");
            processWatch = Stopwatch.StartNew();
            var splitSheet = ReportFile.GetWorkbook(ReportFile.SplitPaymentINGPL).Workbook.Worksheets[0];
            OpeningBalaceSplitING.InsertOpeningBalanceSplitING(MasterData.reportSheet, splitSheet, "SPLIT");
            OpeningBalaceSplitING.InsertOpeningBalanceSplitINGEscrow(MasterData.reportSheet, splitSheet, "Escrow SPLIT");
            processWatch.Stop();
            processMessage = "Success! Split payment opening balances loaded.";
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");

            Console.WriteLine($"Step 6 - Reading Outflows / Inflows....");
            processWatch = Stopwatch.StartNew();
            var outflowInflowSheet = ReportFile.GetWorkbook(ReportFile.OutflowsInflowsINGSPP).Workbook.Worksheets[0];
            Outflow.InsertOutflows(MasterData.reportSheet, outflowInflowSheet, "ING CP");
            Inflow.InsertInflowsING(MasterData.reportSheet, outflowInflowSheet, "ING CP");
            outflowInflowSheet = ReportFile.GetWorkbook(ReportFile.OutflowsInflowsINGCDE).Workbook.Worksheets[0];
            Outflow.InsertOutflows(MasterData.reportSheet, outflowInflowSheet, "ING CP");
            Inflow.InsertInflowsING(MasterData.reportSheet, outflowInflowSheet, "ING CP");
            outflowInflowSheet = ReportFile.GetWorkbook(ReportFile.OutflowsInflowsINGTFI).Workbook.Worksheets[0];
            Outflow.InsertOutflows(MasterData.reportSheet, outflowInflowSheet, "ING N");
            //TODO V1 Do przemyślenia kodowanie bankaccount bo opieranie sie na jednym (ING N / ING CP) nie starcza.
            Inflow.InsertInflowsING(MasterData.reportSheet, outflowInflowSheet, "ING CP");
            processWatch.Stop();
            processMessage = "Success! Outflows and inflows loaded.";
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");

            Console.WriteLine($"Step 7 - Reading daily limits....");
            processWatch = Stopwatch.StartNew();
            var dailyLimitsSheet = ReportFile.GetWorkbook(ReportFile.DailyLimitINGPLEUR).Workbook.Worksheets[0];
            DailyLimit.InsertDailyLimitING(MasterData.reportSheet, dailyLimitsSheet, "ING CP");
            dailyLimitsSheet = ReportFile.GetWorkbook(ReportFile.DailyLimitINGPLPLN).Workbook.Worksheets[0];
            DailyLimit.InsertDailyLimitING(MasterData.reportSheet, dailyLimitsSheet, "ING CP");
            processWatch.Stop();
            processMessage = "Success! Daily limits loaded.";
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");

            Console.WriteLine($"Step 8 - Reading SEB data....");
            processWatch = Stopwatch.StartNew();
            OpeningBalanceSEB.InsertDataSEB(MasterData.reportSheet, "SEB");

            // TODO wykonać jeszcze procedurę dla SEB, gdzie jeżeli w kolumnie N są kwoty,
            // to należy je wpisać do OutFlow i InFlow do arkusza z dnia poprzedniego

            processWatch.Stop();
            processMessage = "Success! Data from SEB loaded.";
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");

            Console.WriteLine($"Step 9 - Reading Santander data....");
            processWatch = Stopwatch.StartNew();
            OpeningBalanceSantander.InsertDataSantander(MasterData.reportSheet, "Santander");

            // TODO wykonać jeszcze procedurę dla Santander, gdzie jeżeli w kolumnie H są kwoty,
            // to należy je wpisać do OutFlow i InFlow do arkusza z dnia poprzedniego

            processWatch.Stop();
            processMessage = "Success! Data from Santander loaded.";
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");

            Console.WriteLine($"Step 10 - Reading ING BV data....");
            processWatch = Stopwatch.StartNew();
            OpeningBalanceINGBV.InsertDataINGBV(MasterData.reportSheet, "ING BV");

            // TODO wykonać jeszcze procedurę dla Santander, gdzie jeżeli w kolumnie H są kwoty,
            // to należy je wpisać do OutFlow i InFlow do arkusza z dnia poprzedniego

            processWatch.Stop();
            processMessage = "Success! Data from ING BV loaded.";
            Console.WriteLine($"\tElapsed: {TimeSpan.FromMilliseconds(processWatch.ElapsedMilliseconds).TotalSeconds} sek.");
            Console.WriteLine($"\tResult: {processMessage}\n");


            MasterData.reportWorkbook.Save();

            Console.WriteLine($"\tTotal ROBOT execution time: {TimeSpan.FromMilliseconds(robotWatch.ElapsedMilliseconds).TotalSeconds} sek.");

            Console.WriteLine("Press any key to finish.");
        }

        public static bool CanCreateReportSheet()
        {
            if (Utils.CheckIfSheetExists(MasterData.reportWorkbook, MasterData.reportDate))
            {
                Console.WriteLine($"Report for {MasterData.reportDate} already exist. Delete it and create new one [y/n]?");
                if (Console.ReadKey().Key != ConsoleKey.Y)
                {
                    return false;
                }
                else
                {
                    MasterData.reportWorkbook.Workbook.Worksheets.Delete(MasterData.reportDate);
                    MasterData.reportWorkbook.Save();
                }
            }
            return true;
        }
    }
}
