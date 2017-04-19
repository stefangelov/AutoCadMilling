using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;

namespace MillingDataEngine.Func
{
    public class ExcelDataRead
    {
        static string ReturnSelectedFilePath()
        {
            string theFilePath = "";
            OpenFileDialog fbd = new OpenFileDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                theFilePath = fbd.FileName;
            }
            return theFilePath;
        }

        public static string[,] ReadData()
        {
            Excel.Application xlApp = new Excel.Application();
            string fileName = ReturnSelectedFilePath();
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(path);
            Excel.Worksheet xlWorkSheet = xlWorkBook.ActiveSheet;
            Excel.Range range = xlWorkSheet.UsedRange;

            int row = range.Rows.Count;
            int col = 14; // this is the number of colums in our case

            string[,] tempExcelDataVariable = new string[row - 2, col + 1];


            for (int i = 0; i < row - 2; i++)
            {
                int j = 0;
                for (; j < col; j++)
                {
                    var tempVariable = (range.Cells[i + 3, j + 1] as Excel.Range).Value;

                    if (tempVariable == null) // if temp variable is null == end of document
                    {
                        row = i;
                        break;
                    }
                    else
                    {
                        tempExcelDataVariable[i, j] = "" + tempVariable;
                    }
                }
                var tempVarSecond = (range.Cells[i + 3, j + 2] as Excel.Range).Value;
                tempExcelDataVariable[i, j] = "" + tempVarSecond;
            }

            xlWorkBook.Close(0);
            xlApp.Quit();

            return tempExcelDataVariable;
        }

        // Create Milling Elements from string Arr from Excel file
        public static DataStruct.RoadSection RoadSectionElementsBuilder(string[,] excelDataVariable)
        {
            DataStruct.RoadSection roadSection = new DataStruct.RoadSection();

            // find dimensions of input arr
            int row = excelDataVariable.GetLength(0);
            int col = excelDataVariable.GetLength(1) - 1;

            for (int rows = 0; rows < row; rows++)
            {
                string[] singleRow = new string[col + 1]; //temp variable to hold single row of input data; + 1 because of min milling depth acordint to Geomechanic rec
                int columns = 0;
                for (; columns < col; columns++)
                {
                    singleRow[columns] = excelDataVariable[rows, columns]; // extract single row from input data to temp variable
                }
                singleRow[columns] = excelDataVariable[rows, columns];
                DataStruct.Cross_section elementsFromSingleRow = ExtractMillingElementsFromSingleRow(singleRow);
                roadSection.AddCross(elementsFromSingleRow);
            }
            return roadSection;
        }

        private static DataStruct.Cross_section ExtractMillingElementsFromSingleRow(string[] singleRow)
        {
            List<DataStruct.MillingElement> singleRowMillingElements = new List<DataStruct.MillingElement>();

            int rowLength = singleRow.GetLength(0) - 1; // - 1 to use old code and take min milling depth from geomechanic
            double station = Convert.ToDouble(singleRow[0]);

            string profilName = singleRow[1];

            int iterationEnd = (rowLength - 4) / 2 + 2 - 1;
            double crossSectionWidth = Convert.ToDouble(singleRow[rowLength - 2]);
            double elementWidth = crossSectionWidth / (iterationEnd - 2); // distance between two points wiht elevation
            double projLayerThick = Convert.ToDouble(singleRow[rowLength - 1]); // Thickness of project asphalt layers
            double allowableMinMillingDepth = Convert.ToDouble(singleRow[rowLength]); // Min alloable milling depth according to geomechanic

            double leftEdgeProjectLevel = 0;
            double midProjectLevel = 0;
            double rightProjectLevel = 0;

            for (int i = 2; i < iterationEnd; i++)
            {
                double existStartLevel;
                double projStartLevel; 
                double existEndLevel;
                double projEndLevel;
                existStartLevel = Convert.ToDouble(singleRow[i]);
                projStartLevel = Convert.ToDouble(singleRow[i + 5]);
                existEndLevel = Convert.ToDouble(singleRow[i + 1]);
                projEndLevel = Convert.ToDouble(singleRow[i + 6]);

                double elementStart = elementWidth * (i - 2) * (-1);

                List<DataStruct.MillingElement> tempListMillingElems = ConvertToMillingElemets(existStartLevel, projStartLevel,
                    existEndLevel, projEndLevel, elementWidth, projLayerThick, station,
                    profilName, elementStart, allowableMinMillingDepth);

                foreach (var item in tempListMillingElems)
                {
                    singleRowMillingElements.Add(item);
                }

                if (i == 2)
                {
                    leftEdgeProjectLevel = projStartLevel;
                }
                if (i == 4)
                {
                    midProjectLevel = projStartLevel;
                }
                if (i == 5)
                {
                    rightProjectLevel = projEndLevel;
                }
            }

            // Element to return
            DataStruct.Cross_section tempCrossSection = new DataStruct.Cross_section(profilName, station, leftEdgeProjectLevel,
                rightProjectLevel, midProjectLevel, crossSectionWidth, singleRowMillingElements);
            tempCrossSection.ProjLayerThick = projLayerThick;

            return tempCrossSection;
        }

        public static double FindNewMillingDepth(double alloableMinMillingDepth, double projLayerThick, double projLevelDiff)
        {
            double firstBinderCourseMin = 6;
            double firstBinderCourseMax = 10;
            double secondBinderCourseMin = 6;
            double secondBinderCourseMax = 20;

            double twoLayerUpLayer = 5;
            double singleLayer = 6;

            double toleranceForMilling = 0.7; //if actual milling is 0.7 * alloableMinMillingDepth everything is OK

            double millingDepth = 0;

            if (projLayerThick < 11)
            {
                if ((alloableMinMillingDepth * toleranceForMilling + projLevelDiff) < projLayerThick)
                {
                    millingDepth = projLayerThick - projLevelDiff;
                }
                else
                {
                    if ((alloableMinMillingDepth + projLevelDiff) < (projLayerThick + firstBinderCourseMin))
                    {
                        millingDepth = (projLayerThick + firstBinderCourseMin) - projLevelDiff;
                    }
                    else
                    {
                        if ((alloableMinMillingDepth + projLevelDiff) < (projLayerThick + firstBinderCourseMin + secondBinderCourseMin))
                        {
                            millingDepth = (projLayerThick + firstBinderCourseMin + secondBinderCourseMin) - projLevelDiff;
                        }
                        else
                        {
                            millingDepth = alloableMinMillingDepth;
                        }
                    }
                }
            }
            else
            {
                if ( (alloableMinMillingDepth + projLevelDiff) < (projLayerThick + firstBinderCourseMax - firstBinderCourseMin) )
                {
                    millingDepth = alloableMinMillingDepth;
                }
                else
                {
                    if ((alloableMinMillingDepth * toleranceForMilling + projLevelDiff) < (projLayerThick + firstBinderCourseMax - firstBinderCourseMin))
                    {
                        millingDepth = projLayerThick + firstBinderCourseMax - firstBinderCourseMin - projLevelDiff;
                    }
                    else
                    {                    
                        if ( (alloableMinMillingDepth + projLevelDiff) < (projLayerThick + secondBinderCourseMin))
                        {
                            millingDepth = projLayerThick + secondBinderCourseMin - projLevelDiff;
                        }
                        else
                        {
                            millingDepth = alloableMinMillingDepth;
                        }
                    }
                }
            }


            return millingDepth;
        }

        // convert secton of cross section to milling elements
        private static List<DataStruct.MillingElement> ConvertToMillingElemets(double existStartLevel, double projStartLevel,
            double existEndLevel, double projEndLevel, double elementWidth, double projLayerThick, double station,
            string profilName, double elementStart, double allowableMinMillingDepth)
        {
            List<DataStruct.MillingElement> listToReturn = new List<DataStruct.MillingElement>();

            double startMillingDepth = (projStartLevel - projLayerThick / 100 - existStartLevel) * -100;
            double endMillingDepth = (projEndLevel - projLayerThick / 100 - existEndLevel) * -100;

            if (startMillingDepth < allowableMinMillingDepth)
            {
                double projLevelDiff = (projStartLevel - existStartLevel) * 100;
                startMillingDepth = FindNewMillingDepth(allowableMinMillingDepth, projLayerThick, projLevelDiff);
            }
            if (endMillingDepth < allowableMinMillingDepth)
            {
                double projLevelDiff = (projEndLevel - existEndLevel) * 100;
                endMillingDepth = FindNewMillingDepth(allowableMinMillingDepth, projLayerThick, projLevelDiff);
            }

            // hold is the milling elemen is in one range of milling depth
            bool areInOneRange = false;
            // check upper and asign to variable
            areInOneRange = (startMillingDepth > DataStruct.MillingElement.MillingRange_1[0] &&
                startMillingDepth <= DataStruct.MillingElement.MillingRange_1[1] &&
                endMillingDepth > DataStruct.MillingElement.MillingRange_1[0] &&
                endMillingDepth <= DataStruct.MillingElement.MillingRange_1[1]) ||
                (startMillingDepth > DataStruct.MillingElement.MillingRange_2[0] &&
                startMillingDepth <= DataStruct.MillingElement.MillingRange_2[1] &&
                endMillingDepth > DataStruct.MillingElement.MillingRange_2[0] &&
                endMillingDepth <= DataStruct.MillingElement.MillingRange_2[1]) ||
                (startMillingDepth > DataStruct.MillingElement.MillingRange_3[0] &&
                startMillingDepth <= DataStruct.MillingElement.MillingRange_3[1] &&
                endMillingDepth > DataStruct.MillingElement.MillingRange_3[0] &&
                endMillingDepth <= DataStruct.MillingElement.MillingRange_3[1]) ||
                (startMillingDepth > DataStruct.MillingElement.MillingRange_3[1] &&
                endMillingDepth > DataStruct.MillingElement.MillingRange_3[1]);

            bool isThereMilling = startMillingDepth >= 0 || endMillingDepth >= 0;
            bool isMillingUnderLastRange = startMillingDepth > DataStruct.MillingElement.MillingRange_3[1] &&
                endMillingDepth > DataStruct.MillingElement.MillingRange_3[1];

            if (areInOneRange || isMillingUnderLastRange)
            {
                listToReturn.Add(new DataStruct.MillingElement(station, profilName, elementStart, elementWidth, startMillingDepth, endMillingDepth));
            }
            else
            {
                if (isThereMilling)
                {
                    double multiplier = (startMillingDepth - endMillingDepth) / elementWidth;
                    List<DataStruct.MillingElement> tempListDifRanges = ConvertMillingElementsInDiffRanges_New(station, profilName, elementStart, elementWidth, startMillingDepth, endMillingDepth, multiplier);

                    if (tempListDifRanges != null)
                    {
                        foreach (var item in tempListDifRanges)
                        {
                            if (item != null)
                            {

                                listToReturn.Add(item);
                            }
                        }
                    }
                }

            }
            return listToReturn;
        }

        private static List<DataStruct.MillingElement> ConvertMillingElementsInDiffRanges_New(double station, string profilName, double elementStart,
            double elementWidth, double startMillingDepth, double endMillingDepth, double multiplier,
            List<DataStruct.MillingElement> theListToReturn = null, int rangeCounter = -1, double[][] theMillingRangess = null)
        {
            if (theListToReturn == null)
            {
                theListToReturn = new List<DataStruct.MillingElement>();
            }

            if (elementWidth <= 0)
            {
                return theListToReturn;
            }

            double[][] theMillingRanges = new double[][] { };

            if (theMillingRangess == null)
            {
                double[] lastPosibleRange = new double[2] { DataStruct.MillingElement.MillingRange_3[1], 100 };

                List<double[]> theMillingRangesList = new List<double[]>();
                foreach (var item in DataStruct.MillingElement.MillingRanges)
                {
                    theMillingRangesList.Add(item);
                }
                theMillingRangesList.Add(lastPosibleRange);
                theMillingRanges = theMillingRangesList.ToArray();
            }
            else
            {
                theMillingRanges = theMillingRangess;
            }

            bool isFirstMillingDepthLarger = startMillingDepth > endMillingDepth;

            if (isFirstMillingDepthLarger)
            {
                if (rangeCounter < 0)
                {
                    rangeCounter = theMillingRanges.Length - 1;
                }
                if (startMillingDepth <= 0)
                {
                    return theListToReturn;
                }
                else
                {
                    if (startMillingDepth < theMillingRanges[rangeCounter][0])
                    {
                        rangeCounter--;
                        ConvertMillingElementsInDiffRanges_New(station, profilName, elementStart, elementWidth,
                            startMillingDepth, endMillingDepth, multiplier, theListToReturn, rangeCounter, theMillingRanges);
                    }
                    else
                    {
                        if (endMillingDepth < theMillingRanges[rangeCounter][0])
                        {
                            double tempStartMilingDepth = startMillingDepth;
                            double tempEndMillingDepth = theMillingRanges[rangeCounter][0];
                            double tempMillingLength = findMillingLenght(tempStartMilingDepth, tempEndMillingDepth, multiplier);
                            theListToReturn.Add(new DataStruct.MillingElement(station, profilName, elementStart, tempMillingLength, tempStartMilingDepth, tempEndMillingDepth));
                            double tempElementStart = elementStart - tempMillingLength;
                            double tempElementWidth = elementWidth - tempMillingLength;
                            rangeCounter--;
                            ConvertMillingElementsInDiffRanges_New(station, profilName,
                                tempElementStart, tempElementWidth, tempEndMillingDepth, endMillingDepth, multiplier, theListToReturn, rangeCounter, theMillingRanges);
                        }
                        else
                        {
                            DataStruct.MillingElement tempMillingElement = new DataStruct.MillingElement(station, profilName, elementStart, elementWidth, startMillingDepth, endMillingDepth);
                            theListToReturn.Add(tempMillingElement);
                        }

                    }
                }
            }
            else
            {
                if (rangeCounter < 0)
                {
                    rangeCounter = 0;
                }
                if (startMillingDepth < 0)
                {
                    double tempEndMillingDepth = theMillingRanges[0][0];
                    double tempStartMillingDepth = startMillingDepth;
                    double tempMillingLength = findMillingLenght(tempStartMillingDepth, tempEndMillingDepth, multiplier);
                    double tempElementSart = elementStart - tempMillingLength;
                    double tempElementWidt = elementWidth - tempMillingLength;
                    ConvertMillingElementsInDiffRanges_New(station, profilName,
                        tempElementSart, tempElementWidt, tempEndMillingDepth, endMillingDepth, multiplier, theListToReturn, rangeCounter, theMillingRanges);
                }
                else
                {
                    if (startMillingDepth > theMillingRanges[rangeCounter][1])
                    {
                        rangeCounter++;
                        ConvertMillingElementsInDiffRanges_New(station, profilName,
                            elementStart, elementWidth, startMillingDepth, endMillingDepth, multiplier, theListToReturn, rangeCounter, theMillingRanges);
                    }
                    else
                    {
                        if (endMillingDepth > theMillingRanges[rangeCounter][1])
                        {
                            double tempEndMillingDepth = theMillingRanges[rangeCounter][1];
                            double tempStartMilingDepth = startMillingDepth;
                            double tempMillingLength = findMillingLenght(tempStartMilingDepth, tempEndMillingDepth, multiplier);
                            theListToReturn.Add(new DataStruct.MillingElement(station, profilName, elementStart, tempMillingLength, tempStartMilingDepth, tempEndMillingDepth));
                            double tempElementStart = elementStart - tempMillingLength;
                            double tempElementWidth = elementWidth - tempMillingLength;
                            rangeCounter++;
                            ConvertMillingElementsInDiffRanges_New(station, profilName, tempElementStart,
                                tempElementWidth, tempEndMillingDepth, endMillingDepth, multiplier, theListToReturn, rangeCounter, theMillingRanges);
                        }
                        else
                        {
                            theListToReturn.Add(new DataStruct.MillingElement(station, profilName, elementStart, elementWidth, startMillingDepth, endMillingDepth));
                            return theListToReturn;
                        }
                    }
                }
            }
            return theListToReturn;
        }
        private static double findMillingLenght(double tempStartMillingDepth, double tempEndMillingDepth, double multiplier)
        {
            double lengthToReturn = Math.Abs((tempStartMillingDepth - tempEndMillingDepth) / multiplier);
            return lengthToReturn;
        }
    }
}