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
                foreach (var item in cross.MillingElements)
                {
                    readymillingelements.Add(item);
                }

            }
            Console.WriteLine("done!");

            int tempcounter = 0;

            foreach (var item in readymillingelements)
            {
                if (tempcounter % 10 == 0)
                {
                    Console.WriteLine();
                }
                Console.WriteLine("station {0}, point {1}, layer: {2}", item.Station, item.ProfileName, item.LayerName);
                Console.WriteLine("start point: {0} - {1}", item.StartPoint.CoordinateX, item.StartPoint.CoordinateY);
                Console.WriteLine("end point: {0} - {1}", item.EndPoint.CoordinateX, item.EndPoint.CoordinateY);
                Console.WriteLine();
                             
                tempcounter++;
            }
        }
    }
}
