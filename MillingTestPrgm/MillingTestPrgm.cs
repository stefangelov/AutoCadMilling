using System;
using System.Collections.Generic;

namespace MillingTestPrgm
{
    class AutoCad_milling
    {
        // use this to test withoud open Acad or C3D
        [STAThread]
        static void Main(string[] args)
        {
            //double baseElevation = MillingDataEngine.Func.SectionViews.ElevationOfLocationPoint(1192.29915019433);
            //Console.WriteLine(baseElevation);

            string[,] exceldatavariable = MillingDataEngine.Func.ExcelDataRead.ReadData();

            List<MillingDataEngine.DataStruct.Cross_section> thecrosssection = MillingDataEngine.Func.ExcelDataRead.RoadSectionElementsBuilder(exceldatavariable).CrossSections;
            List<MillingDataEngine.DataStruct.MillingElement> readymillingelements = new List<MillingDataEngine.DataStruct.MillingElement>();
            foreach (var cross in thecrosssection)
            {
                if (cross.MillingElements.Count > 0)
                {
                    Console.WriteLine("station {0}, Nmae: {1}", cross.MillingElements[0].Station, cross.MillingElements[0].ProfileName);
                    foreach (var quont in cross.MillingQuantity)
                    {
                        string forPrint = (quont.Range[1]) > 0 ? 
                            String.Format("\tRange {0}-{1}  Q:{2}", quont.Range[0], quont.Range[1], quont.Quant) : 
                            String.Format("\tRange >{0}  Q:{1}", quont.Range[0], quont.Quant);
                        Console.WriteLine(forPrint);
                    }
                }
            
            }
                Console.WriteLine("done!");
        }
    }
}
