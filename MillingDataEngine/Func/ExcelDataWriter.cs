using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;

namespace MillingDataEngine.Func
{
    public class ExcelDataWriter
    {
        public static void DisplayInExcel(IEnumerable<MillingDataEngine.DataStruct.Cross_section> crossSections)
        {
            var excelApp = new Excel.Application();
            // Make the object visible.
            excelApp.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            excelApp.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            // Establish column headings in cells A1 and B1.
            workSheet.Cells[1, "A"] = "Station";
            workSheet.Cells[1, "B"] = "Profile Name";
            workSheet.Cells[1, "C"] = "str_0-3";
            workSheet.Cells[1, "E"] = "str_3-5";
            workSheet.Cells[1, "G"] = "str_5-7";
            workSheet.Cells[1, "I"] = "str_7-";

            var row = 1;
            foreach (var cs in crossSections)
            {
                row++;

                workSheet.Cells[row, "A"] = cs.Station;
                workSheet.Cells[row, "B"] = cs.Name;

                foreach (var quant in cs.MillingQuantity)
                {
                    if (quant.Range[0] == 0 && quant.Range[1] == 3)
                    {
                        workSheet.Cells[row, "C"] = String.Format("{0:0.00}", quant.Quant);
                    }
                    else
                    {
                        if (quant.Range[0] == 3 && quant.Range[1] == 5)
                        {
                            workSheet.Cells[row, "E"] = String.Format("{0:0.00}", quant.Quant);
                        }
                        else
                        {
                            if (quant.Range[0] == 5 && quant.Range[1] == 7)
                            {
                                workSheet.Cells[row, "G"] = String.Format("{0:0.00}", quant.Quant);
                            }
                            else
                            {
                                if (quant.Range[0] >= 7)
                                {
                                    workSheet.Cells[row, "I"] = String.Format("{0:0.00}", quant.Quant);
                                }
                            }
                        }
                    }
                }
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();

            //to fit columns according to data in them
            for (int i = 1; i < 13; i++)
            {
                ((Excel.Range)workSheet.Columns[i]).AutoFit();
            }

            excelApp.Quit();
        }
    }
}
