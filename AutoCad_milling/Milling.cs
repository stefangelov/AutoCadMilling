using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System.Collections.Generic;
using System;
using Autodesk.Civil.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.Civil.DatabaseServices;

namespace AutoCad_milling
{
    public class Milling
    {
        // the name of Acad Command to insert milling depth in plan view
        [CommandMethod("MillingDiagramCreator")]
        [STAThread]
        public static void CreateAddMillingElements()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;


            // Start a transaction to create necessary layers
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;

                // Add layers in layer table
                MillingDataEngine.Func.Layers.AddLayers("StA_Strugane_7", 1, acTrans, acLyrTbl);
                MillingDataEngine.Func.Layers.AddLayers("StA_Strugane_5_7", 5, acTrans, acLyrTbl);
                MillingDataEngine.Func.Layers.AddLayers("StA_Strugane_3_5", 4, acTrans, acLyrTbl);
                MillingDataEngine.Func.Layers.AddLayers("StA_Strugane_0_3", 2, acTrans, acLyrTbl);
                MillingDataEngine.Func.Layers.AddLayers("Border", 1, acTrans, acLyrTbl);
                MillingDataEngine.Func.Layers.AddLayers("Axis", 1, acTrans, acLyrTbl); // After drawing creation You must change line type manual.
                MillingDataEngine.Func.Layers.AddLayers("Texts", 7, acTrans, acLyrTbl);

                
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;

                // create test Elements (Next change the name from "Test" to something other)
                MillingDataEngine.DataStruct.RoadSection theRoadSection = MillingDataEngine.Func.ExcelDataRead.RoadSectionElementsBuilder(MillingDataEngine.Func.ExcelDataRead.ReadData());
                List<MillingDataEngine.DataStruct.MillingElement> listOfMillingElements = new List<MillingDataEngine.DataStruct.MillingElement>();
                
                foreach (var section in theRoadSection.CrossSections)
                {
                    foreach (var item in section.MillingElements)
                    {
                        listOfMillingElements.Add(item);
                    }
                }

                MillingDataEngine.DataStruct.MillingElement[] allTestElements = listOfMillingElements.ToArray();
                // drow test Elements
                MillingDataEngine.Func.DrawMillingElements.Drow(allTestElements, acTrans, acBlkTblRec, acCurDb);
                
                // Save the changes and dispose of the transaction
                acTrans.Commit();
            }
        }

        // the name of Acad Command to insert milling depth in Profile View
        [CommandMethod ("MillingToProfileView")]
        [STAThread]
        public static void CreateMillingElementsToProfileView()
        {
            CivilDocument doc = CivilApplication.ActiveDocument;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            using (Transaction ts = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
            {
                var promptResultsPV = ed.GetEntity("Selec Profile View");
                
            }
        }

    }
}
