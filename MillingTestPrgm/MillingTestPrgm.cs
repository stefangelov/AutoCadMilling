using System;
using System.Collections.Generic;

namespace MillingTestPrgm
{
    class AutoCad_milling
    {

         public class Account
    {
        public int ID { get; set; }
        public double Balance { get; set; }
    }

        // use this to test withoud open Acad or C3D
        [STAThread]
        static void Main(string[] args)
        {

            string[,] exceldatavariable = MillingDataEngine.Func.ExcelDataRead.ReadData();

            List<MillingDataEngine.DataStruct.Cross_section> crosssections = MillingDataEngine.Func.ExcelDataRead.RoadSectionElementsBuilder(exceldatavariable).CrossSections;

            MillingDataEngine.Func.ExcelDataWriter.DisplayInExcel(crosssections);

            
            Console.WriteLine("done!");
        }
    }
}
