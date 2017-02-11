using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System;
using MillingDataEngine.DataStruct;

namespace MillingDataEngine.Func
{
    public class DrawMillingElements
    {
        public static void Drow(MillingDataEngine.DataStruct.MillingElement[] allMillingElements, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb)
        {
            string tempProfileName = ""; //hold name of Profile to chech if we pass to different profile
            string tempMillingElementProfilName = allMillingElements[0].ProfileName;
            int millingElementIndex = 0;
            int millingElementAllIndexes = allMillingElements.Length;
            foreach (var millingElement in allMillingElements)
            {
                // add lebel for name
                // add label for station
                if (tempProfileName != millingElement.ProfileName)
                {
                    // LabelingProfileName Profile Names
                    LabelingProfileName(millingElement, acTrans, acBlkTblRec, acCurDb);

                    // LabelingProfileStation Profile Stations
                    LabelingProfileStation(millingElement, acTrans, acBlkTblRec, acCurDb);

                    tempProfileName = millingElement.ProfileName;
                }

                // add miling depth on end of cross section
                if (millingElement.ProfileName != tempMillingElementProfilName || millingElementIndex == millingElementAllIndexes - 1)
                {
                    LabelingMillingEndDepth(allMillingElements[millingElementIndex - 1], acTrans, acBlkTblRec, acCurDb);
                    tempMillingElementProfilName = millingElement.ProfileName;
                }

                // set necessary layer
                MillingDataEngine.Func.Layers.SetAsCurrent(millingElement.LayerName, acCurDb);

                // Create a line 
                Line acLine = new Line(new Point3d(millingElement.StartPoint.CoordinateX, millingElement.StartPoint.CoordinateY, millingElement.StartPoint.CoordinateZ),
                new Point3d(millingElement.EndPoint.CoordinateX, millingElement.EndPoint.CoordinateY, millingElement.EndPoint.CoordinateZ));
                acLine.SetDatabaseDefaults();

                // Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acLine);
                acTrans.AddNewlyCreatedDBObject(acLine, true);

                // LabelingMillingDepth Milling Items
                LabelingMillingDepth(millingElement, acTrans, acBlkTblRec, acCurDb);

                millingElementIndex++;
            }
        }

        private static void LabelingMillingEndDepth(DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb)
        {
            double deltaRefTextX = 0.35;
            // Open the Block table for read
            BlockTable acBlkTbl;
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
            OpenMode.ForRead) as BlockTable;

            // Open the Block table record Model space for write
            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

            // set necessary layer
            MillingDataEngine.Func.Layers.SetAsCurrent(millingElement.LayerName, acCurDb);

            // Create a single-line text object for milling depth
            DBText acText = new DBText();
            acText.SetDatabaseDefaults();
            acText.Position = new Point3d(millingElement.EndPoint.CoordinateX + deltaRefTextX, millingElement.EndPoint.CoordinateY, millingElement.EndPoint.CoordinateZ);
            acText.Height = 0.25;
            int roundedMillingDepth = Convert.ToInt32(Math.Round(millingElement.EndMillingDepth, 0));
            if (roundedMillingDepth == 0)
            {
                acText.TextString = roundedMillingDepth + " cm";
            }
            else
            {
                acText.TextString = "-" + roundedMillingDepth + " cm";
            }

            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }

        // Labeling profil names
        private static void LabelingProfileName(DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb)
        {
            double deltaRefTextX = -2.5;
            // Open the Block table for read
            BlockTable acBlkTbl;
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
            OpenMode.ForRead) as BlockTable;

            // Open the Block table record Model space for write
            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

            // set necessary layer
            MillingDataEngine.Func.Layers.SetAsCurrent("Texts", acCurDb);

            // Create a single-line text object for Profile name
            DBText acText = new DBText();
            acText.SetDatabaseDefaults();
            acText.Position = new Point3d(millingElement.StartPoint.CoordinateX + deltaRefTextX, millingElement.StartPoint.CoordinateY, millingElement.StartPoint.CoordinateZ);
            acText.Height = 0.5;
            acText.TextString = millingElement.ProfileName;

            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }

        // Labeling profil stations
        private static void LabelingProfileStation(DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb)
        {
            double deltaRefTextX = -2.5;
            double deltaRefTextY = 2.5;
            // Open the Block table for read
            BlockTable acBlkTbl;
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
            OpenMode.ForRead) as BlockTable;

            // Open the Block table record Model space for write
            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

            // set necessary layer
            MillingDataEngine.Func.Layers.SetAsCurrent("Texts", acCurDb);

            // Create a single-line text object for Profile name
            DBText acText = new DBText();
            acText.SetDatabaseDefaults();
            acText.Position = new Point3d(millingElement.StartPoint.CoordinateX + deltaRefTextX, millingElement.StartPoint.CoordinateY + deltaRefTextY, millingElement.StartPoint.CoordinateZ);
            acText.Height = 0.5;

            string stringToPut = millingElement.Station.ToString();
            int stringToPutLength = stringToPut.Length;
            if (stringToPutLength > 3)
            {
                stringToPut = stringToPut.Remove(stringToPutLength - 3) + '+' + stringToPut.Substring(stringToPutLength - 3);
            }
            else
            {
                if (stringToPutLength == 3)
                {
                    stringToPut = "0+" + stringToPut;
                }
                else
                {
                    if (stringToPutLength == 2)
                    {
                        stringToPut = "0+0" + stringToPut;
                    }
                    else
                    {
                        stringToPut = "0+00" + stringToPut;
                    }
                }
            }

            acText.TextString = stringToPut;

            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }

        // labeling milling depth on crossectons
        private static void LabelingMillingDepth(MillingDataEngine.DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb)
        {
            double deltaRefTextX = 0.35;
            // Open the Block table for read
            BlockTable acBlkTbl;
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
            OpenMode.ForRead) as BlockTable;
         
            // Open the Block table record Model space for write
            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

            // set necessary layer
            MillingDataEngine.Func.Layers.SetAsCurrent(millingElement.LayerName, acCurDb);

            // Create a single-line text object for milling depth
            DBText acText = new DBText();
            acText.SetDatabaseDefaults();
            acText.Position = new Point3d(millingElement.StartPoint.CoordinateX + deltaRefTextX, millingElement.StartPoint.CoordinateY, millingElement.StartPoint.CoordinateZ);
            acText.Height = 0.25;
            int roundedMillingDepth = Convert.ToInt32(Math.Round(millingElement.StartMillingDepth, 0));
            if (roundedMillingDepth == 0)
            {
                acText.TextString = roundedMillingDepth + " cm";
            }
            else
            {
                acText.TextString = "-" + roundedMillingDepth + " cm";
            }

            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }
    }
}
