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
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(path); // on other machine change path to wor proper
            Excel.Worksheet xlWorkSheet = xlWorkBook.ActiveSheet;
            Excel.Range range = xlWorkSheet.UsedRange;

            int row = range.Rows.Count;
            int col = 14; // this is the number of colums in our case

            string[,] tempExcelDataVariable = new string [row - 2, col];


            for (int i = 0; i < row - 2; i++)
            {
                for (int j = 0; j < col; j++)
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
            }

            xlWorkBook.Close(0);
            xlApp.Quit();

            string [,] excelDataVariable = new string [row, col];

            for (int i = 0; i < row; i++)
            {
                for (int j = 0; j < col; j++)
                {
                    excelDataVariable[i, j] = tempExcelDataVariable[i, j];
                }
            }
            return excelDataVariable;
        }

        // Create Milling Elements from string Arr from Excel file
        public static DataStruct.RoadSection RoadSectionElementsBuilder(string[,] excelDataVariable)
        {
            DataStruct.RoadSection roadSection = new DataStruct.RoadSection();

            // find dimensions of input arr
            int row = excelDataVariable.GetLength(0);
            int col = excelDataVariable.GetLength(1);

            for (int rows = 0; rows < row; rows++)
            {
                string[] singleRow = new string[col]; //temp variable to hold single row of input data

                for (int columns = 0; columns < col; columns++)
                {
                    singleRow[columns] = excelDataVariable[rows, columns]; // extract single row from input data to temp variable
                }
                
                DataStruct.Cross_section elementsFromSingleRow = ExtractMillingElementsFromSingleRow(singleRow);
                roadSection.AddCross(elementsFromSingleRow);
            }
            return roadSection;
        }

        private static DataStruct.Cross_section ExtractMillingElementsFromSingleRow(string[] singleRow)
        {
            List<DataStruct.MillingElement> singleRowMillingElements = new List<DataStruct.MillingElement>();

            int rowLength = singleRow.GetLength(0);
            double station = Convert.ToDouble(singleRow[0]);
            string profilName = singleRow[1];
            int iterationEnd = (rowLength - 4) / 2 + 2 - 1;
            double crossSectionWidth = Convert.ToDouble(singleRow[rowLength - 2]);
            double elementWidth = crossSectionWidth / (iterationEnd - 2); // distance between two points wiht elevation
            double projLayerThick = Convert.ToDouble(singleRow[rowLength - 1]); // Thickness of project asphalt layers

            double leftEdgeProjectLevel = 0;
            double midProjectLevel = 0;
            double rightProjectLevel = 0;

            for (int i = 2; i < iterationEnd; i++)
            {
                double existStartLevel = Convert.ToDouble(singleRow[i]);
                double projStartLevel = Convert.ToDouble(singleRow[i+5]);
                double existEndLevel = Convert.ToDouble(singleRow[i+1]);
                double projEndLevel = Convert.ToDouble(singleRow[i+6]);

                double elementStart = elementWidth * (i - 2) * (-1);

                List<DataStruct.MillingElement>  tempListMillingElems = ConvertToMillingElemets(existStartLevel, projStartLevel,
                    existEndLevel, projEndLevel, elementWidth, projLayerThick, station, 
                    profilName, elementStart);

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

            return tempCrossSection;
        }

        // convert secton of cross section to milling elements
        private static List<DataStruct.MillingElement> ConvertToMillingElemets(double existStartLevel, double projStartLevel, 
            double existEndLevel, double projEndLevel, double elementWidth, double projLayerThick, double station,
            string profilName, double elementStart)
        {
            List<DataStruct.MillingElement> listToReturn = new List<DataStruct.MillingElement>();

            double startMillingDepth = (projStartLevel - projLayerThick / 100 - existStartLevel) * -100;
            double endMillingDepth = (projEndLevel - projLayerThick / 100 - existEndLevel) * -100;

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

            if (areInOneRange)
            {
                listToReturn.Add(new DataStruct.MillingElement(station, profilName, elementStart, elementWidth, startMillingDepth, endMillingDepth));
            }
            else
            {
                List<DataStruct.MillingElement> tempListDifRanges = ConvertToMillingElementsInDifferentRanges(station, profilName, elementStart, 
                    elementWidth, startMillingDepth, endMillingDepth);

                foreach (var item in tempListDifRanges)
                {
                    if (item != null)
                    {
                        listToReturn.Add(item);
                    }
                }
            }

            return listToReturn;
        }

        private static List<DataStruct.MillingElement> ConvertToMillingElementsInDifferentRanges(double station, string profilName, double elementStart, 
            double elementWidth, double startMillingDepth, double endMillingDepth)
        {
            List<DataStruct.MillingElement> listToReturnDifRanges = new List<DataStruct.MillingElement>();

            // to hold millin Range where are start and end of milling line
            double[] startMillingRange = new double[2];
            double[] endMillingRange = new double[2];

            double[][] millingRanges = DataStruct.MillingElement.MillingRanges;
            double[] lastPosibleRange = millingRanges[millingRanges.Length - 1];
            double[] firstPosibleRange = millingRanges[0];

            // search in wich milling ranges are the start ana edn milling depth
            foreach (var range in millingRanges)
            {
                if (startMillingDepth >= range[0] && startMillingDepth < range[1])
                {
                    startMillingRange = range;
                }

                if (endMillingDepth >= range[0] && endMillingDepth < range[1])
                {
                    endMillingRange = range;
                }
            }

            // check if milling element goes through last milling Range
            if (startMillingDepth >= lastPosibleRange[1])
            {
                startMillingRange = new double[] { lastPosibleRange[1], 150 };
            }

            if (endMillingDepth >= lastPosibleRange[1])
            {
                endMillingRange = new double[] { lastPosibleRange[1], 150 };
            }

            // check if milling element goes through first posible range
            if (startMillingDepth < firstPosibleRange[0])
            {
                startMillingRange = new double[] { -100, firstPosibleRange[0] };
            }
            if (endMillingDepth < firstPosibleRange[0])
            {
                endMillingRange = new double[] { -100, firstPosibleRange[0] };
            }

            // check if there is MILLING
            if (startMillingDepth > firstPosibleRange[0] || endMillingDepth > firstPosibleRange[0])
            {
                // chek wich milling depth (start or end) are bigger
                if (startMillingDepth > endMillingDepth)
                {
                    // check if the milling element start before first milling range
                    // with last milling range everything is OK
                    if (startMillingDepth < firstPosibleRange[0])
                    {
                        double[] firstSecondMillingLength = CalcultaFirstSecondMillingLength(elementWidth, startMillingDepth, endMillingDepth, startMillingRange[1], endMillingRange[0]);
                        double unneMillingElementLength = firstSecondMillingLength[0];
                        
                        double neMillingElementStart = elementStart - unneMillingElementLength;
                        double neMillingLength = firstSecondMillingLength[1];

                        listToReturnDifRanges.Add(new DataStruct.MillingElement(station, profilName, neMillingElementStart, neMillingLength, firstPosibleRange[0], endMillingDepth));
                    }
                    else
                    {
                        double[] firstSecondMillingLength = CalcultaFirstSecondMillingLength(elementWidth, startMillingDepth, endMillingDepth, startMillingRange[1], endMillingRange[0]);
                        listToReturnDifRanges.Add(new DataStruct.MillingElement(station, profilName, elementStart, firstSecondMillingLength[0], startMillingDepth, startMillingRange[0]));

                        double secondElementStart = elementStart - firstSecondMillingLength[0];
                        listToReturnDifRanges.Add(new DataStruct.MillingElement(station, profilName, secondElementStart, firstSecondMillingLength[1], startMillingRange[0], endMillingDepth));  
                    }
                }
                else
                {
                    // chech if the milling element start before first milling range
                    // with last milling range everything is OK
                    if (startMillingDepth < firstPosibleRange[0])
                    {
                        double[] neMillingLength = CalcultaFirstSecondMillingLength(elementWidth, startMillingDepth, endMillingDepth, startMillingRange[1], endMillingRange[0]);
                        
                        listToReturnDifRanges.Add(new DataStruct.MillingElement(station, profilName, elementStart + neMillingLength[0], neMillingLength[1], firstPosibleRange[0], endMillingDepth));
                    }

                    else
                    {
                        double[] firstSecondMillingLength = CalcultaFirstSecondMillingLength(elementWidth, startMillingDepth, endMillingDepth, startMillingRange[1], endMillingRange[0]);
                        listToReturnDifRanges.Add(new DataStruct.MillingElement(station, profilName, elementStart, firstSecondMillingLength[0], startMillingDepth, startMillingRange[1]));

                        double secondElementStart = elementStart - firstSecondMillingLength[0];
                        listToReturnDifRanges.Add(new DataStruct.MillingElement(station, profilName, secondElementStart, firstSecondMillingLength[1], startMillingRange[1], endMillingDepth));
                    }
                }
            }
            
            return listToReturnDifRanges;
        }
        
        // Calculate Milling length when milling element goes thrue one milling rage
        private static double[] CalcultaFirstSecondMillingLength(double elementWidth, double startMillingDepth,
            double endMillingDepth, double startMillingRange, double endMillingRange)
        {
            double[] firstSecondMilignLength = new double[2];
            
            double deltaStart = startMillingRange - startMillingDepth; 
            if (deltaStart < 0)
            {
                deltaStart = deltaStart * (-1);
            }

            double deltaEnd = -endMillingRange - endMillingDepth;
            if (deltaEnd < 0)
            {
                deltaEnd = deltaEnd * (-1);
            }

            firstSecondMilignLength[0] = (deltaStart / deltaEnd) * (elementWidth / ((deltaStart / deltaEnd) + 1));
            firstSecondMilignLength[1] = elementWidth - firstSecondMilignLength[0];

            return firstSecondMilignLength;
        }
    }
}