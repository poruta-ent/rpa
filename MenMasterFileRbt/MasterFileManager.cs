using System;
using OfficeOpenXml;

namespace MenMasterFileRbt
{
    public class MenMasterFileManager
    {
        public string MenMasterFilePath => @"C:\Users\szymon.m\Desktop\XXX\";
        public string MenMasterFileName => @"PS Budget 2020_FC04 working.xlsx";
        public string MenMasterFileWorkingFCSheetName => "FC04";


        public static ExcelPackage masterFileBook = new FileInfo(MenMasterFilePath + MenMasterFileName);
        //public static ExcelWorksheet fcSheet = masterFileBook.
        //public static ExcelWorksheet destinationSheet = 

        //TODO: Mange opened file 

        public void PreparePMFile (string[] pms)
        {
            if (pms == null) return;



        }

    }

}