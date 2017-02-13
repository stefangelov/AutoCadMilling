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
            double baseElevation = MillingDataEngine.Func.SectionViews.ElevationOfLocationPoint(1192.29915019433);
            Console.WriteLine(baseElevation);

            //string[,] excelDataVariable = MillingDataEngine.Func.ExcelDataRead.ReadData();

            //List<MillingDataEngine.DataStruct.Cross_section> theCrossSection = MillingDataEngine.Func.ExcelDataRead.RoadSectionElementsBuilder(excelDataVariable).CrossSections;
            //List<MillingDataEngine.DataStruct.MillingElement> readyMillingElements = new List<MillingDataEngine.DataStruct.MillingElement>();
            //foreach (var cross in theCrossSection)
            //{
            //    foreach (var item in cross.MillingElements)
            //    {
            //        readyMillingElements.Add(item);
            //    }
               
            //}
            //Console.WriteLine("Done!");

            //int tempCounter = 0;

            //foreach (var item in readyMillingElements)
            //{
            //    if (tempCounter % 10 == 0)
            //    {
            //        Console.WriteLine();
            //    }
            //    Console.WriteLine("Station {0}, Point {1}, Layer: {2}", item.Station, item.ProfileName, item.LayerName);
            //    Console.WriteLine("Start point: {0} - {1}", item.StartPoint.CoordinateX, item.StartPoint.CoordinateY);
            //    Console.WriteLine("End point: {0} - {1}", item.EndPoint.CoordinateX, item.EndPoint.CoordinateY);
            //    Console.WriteLine();

            //    tempCounter++;
            //}
        }
    }
}
