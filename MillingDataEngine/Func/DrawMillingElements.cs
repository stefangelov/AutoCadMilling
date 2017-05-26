using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using System;
using MillingDataEngine.DataStruct;
using System.Linq;
using System.Collections.Generic;

namespace MillingDataEngine.Func
{
    public class DrawMillingElements
    {
        public static void Drow(MillingDataEngine.DataStruct.RoadSection theRoadSection, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb, bool isToProfileView = false)
        {

            List<MillingDataEngine.DataStruct.MillingElement> listOfMillingElements = new List<MillingDataEngine.DataStruct.MillingElement>();

            if (!isToProfileView)
            {
                foreach (var section in theRoadSection.CrossSections)
                {
                    foreach (var item in section.MillingElements)
                    {
                        listOfMillingElements.Add(item);
                    }
                }
            }
            else
            {
                foreach (var section in theRoadSection.CrossSectionsForSectionView)
                {
                    foreach (var item in section.MillingElements)
                    {
                        listOfMillingElements.Add(item);
                    }
                }
            }

            MillingDataEngine.DataStruct.MillingElement[] allMillingElements = listOfMillingElements.ToArray();

            string tempProfileName = ""; //hold name of Profile to chech if we pass to different profile
            string tempMillingElementProfilName = allMillingElements[0].ProfileName;
            int millingElementIndex = 0;
            int millingElementAllIndexes = allMillingElements.Length;
            foreach (var millingElement in allMillingElements)
            {
                if (!isToProfileView)
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
                    // LabelingMillingDepth Milling Items
                    LabelingMillingDepth(millingElement, acTrans, acBlkTblRec, acCurDb);
                }
                else
                {
                    // add miling depth on end of cross section
                    if (millingElement.ProfileName != tempMillingElementProfilName || millingElementIndex == millingElementAllIndexes - 1)
                    {
                        LabelingMillingEndDepth(allMillingElements[millingElementIndex - 1], acTrans, acBlkTblRec, acCurDb, -.25, 0.18, -0.35);
                        tempMillingElementProfilName = millingElement.ProfileName;
                    }
                    // LabelingMillingDepth Milling Items
                    LabelingMillingDepth(millingElement, acTrans, acBlkTblRec, acCurDb, -.25, 0.18, -0.35);
                }

                // set necessary layer
                MillingDataEngine.Func.Layers.SetAsCurrent(millingElement.LayerName, acCurDb);

                // Create a line
                if (isToProfileView)
                {
                    Line acLine = new Line(new Point3d(millingElement.StartPoint.CoordinateX, millingElement.StartPoint.CoordinateY, millingElement.StartPoint.CoordinateZ),
                    new Point3d(millingElement.EndPoint.CoordinateX, millingElement.EndPoint.CoordinateY, millingElement.EndPoint.CoordinateZ));
                    acLine.SetDatabaseDefaults();
                    // Add the new object to the block table record and the transaction
                    acBlkTblRec.AppendEntity(acLine);
                    acTrans.AddNewlyCreatedDBObject(acLine, true);
                }

                millingElementIndex++;
            }
            if (!isToProfileView)
            {
                LebelingMillingQuantity(theRoadSection, acTrans, acBlkTblRec, acCurDb);
            }
            else
            {
                LebelingMillingQuantity(theRoadSection, acTrans, acBlkTblRec, acCurDb, 1.75, 0.5, 0, 0.18, 0, 0.18, true);
            }
            
        }

        private static void LebelingMillingQuantity(RoadSection theRoadSection, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb,
            double deltaRefTextX = 0, double deltaRefTextY = 0.5, double deltaRefXForEachQuant = 0.18, double deltaRefYForEachQuant = 0, double textRotation = 1.57, double textHight = 0.18,
            bool isForProfileView = false)
        {
            //double deltaRefTextX = 0;
            //double deltaRefTextY = 0.5;
            //double deltaRefXForEachQuant = 0.18;

            // Open the Block table for read
            BlockTable acBlkTbl;
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
            OpenMode.ForRead) as BlockTable;

            // Open the Block table record Model space for write
            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

            foreach (var section in theRoadSection.CrossSections)
            {
                int quantCounter = 0;

                if (section.MillingElements.Count > 0)
                {
                    foreach (var quant in section.MillingQuantity)
                    {
                        // set necessary layer
                        MillingDataEngine.Func.Layers.SetAsCurrent(quant.LayerName, acCurDb);

                        // Create a single-line text object for quant
                        DBText acText = new DBText();
                        acText.SetDatabaseDefaults();
                        acText.Position = new Point3d(section.MillingElements[0].StartPoint.CoordinateX + deltaRefTextX - quantCounter * deltaRefXForEachQuant + (isForProfileView ? section.Width : 0),
                            section.MillingElements[0].StartPoint.CoordinateY + quantCounter * deltaRefYForEachQuant +
                            deltaRefTextY, section.MillingElements[0].StartPoint.CoordinateZ);
                        acText.Height = textHight;

                        string stringToPut = (quant.Range[1]) > 0 ?
                                String.Format("tr{0}_{1}={2:0.00}m", quant.Range[0], quant.Range[1], quant.Quant) :
                                String.Format("str{0}_={1:0.00}m", quant.Range[0], quant.Quant); ;

                        acText.TextString = stringToPut;
                        acText.Rotation = textRotation;
                        acBlkTblRec.AppendEntity(acText);
                        acTrans.AddNewlyCreatedDBObject(acText, true);
                        quantCounter++;
                    }
                    quantCounter = 0;
                }
            }
        }

        private static void LabelingMillingEndDepth(DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb,
            double deltaRefTextX = 0.35, double textHight = 0.25, double deltaYref = 0)
        {
            //double deltaRefTextX = 0.35;
            // Open the Block table for read
            BlockTable acBlkTbl;
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
            OpenMode.ForRead) as BlockTable;

            // Open the Block table record Model space for write
            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

            // set necessary layer
            MillingDataEngine.Func.Layers.SetAsCurrent("MillingDepth", acCurDb);

            // Create a single-line text object for milling depth
            DBText acText = new DBText();
            acText.SetDatabaseDefaults();
            acText.Position = new Point3d(millingElement.EndPoint.CoordinateX + deltaRefTextX, millingElement.EndPoint.CoordinateY + deltaYref, millingElement.EndPoint.CoordinateZ);
            acText.Height = textHight;
            int roundedMillingDepth = Convert.ToInt32(Math.Round(millingElement.EndMillingDepth, 0));
            if (roundedMillingDepth == 0)
            {
                acText.TextString = roundedMillingDepth + "cm";
            }
            else
            {
                acText.TextString = "-" + roundedMillingDepth + "cm";
            }

            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }

        // Labeling profil names
        private static void LabelingProfileName(DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb)
        {
            double deltaRefTextX = -3.0;
            double deltaRefTextY = 0;
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
            acText.Justify = AttachmentPoint.MiddleCenter;
            acText.AlignmentPoint = new Point3d(millingElement.StartPoint.CoordinateX + deltaRefTextX, 50 + deltaRefTextY, millingElement.StartPoint.CoordinateZ);
            acText.Height = 0.5;
            acText.TextString = millingElement.ProfileName;

            acText.Rotation = 1.57;

            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }

        // Labeling profil stations
        private static void LabelingProfileStation(DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb)
        {
            double deltaRefTextX = -2.4;
            double deltaRefTextY = 0;
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
            acText.Justify = AttachmentPoint.MiddleCenter;
            acText.AlignmentPoint = new Point3d(millingElement.StartPoint.CoordinateX + deltaRefTextX, 50 + deltaRefTextY, millingElement.StartPoint.CoordinateZ);
            acText.Height = 0.5;

            string stringToPut = millingElement.Station.ToString();

            if (stringToPut.Contains('.'))
            {
                int stringLength = stringToPut.Length;
                int positionOfDecimal = stringToPut.IndexOf('.');
                string afterDecimal = stringToPut.Substring(positionOfDecimal);
                stringToPut = AddPlusToStation(stringToPut.Substring(0, positionOfDecimal)) + afterDecimal;
            }

            else
            {
                stringToPut = AddPlusToStation(stringToPut);
                stringToPut = stringToPut + ".00";
            }

            acText.TextString = stringToPut;
            acText.Rotation = 1.57;

            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }

        private static string AddPlusToStation(string stringToPut)
        {
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

            return stringToPut;
        }

        // labeling milling depth on crossectons
        private static void LabelingMillingDepth(MillingDataEngine.DataStruct.MillingElement millingElement, Transaction acTrans, BlockTableRecord acBlkTblRec, Database acCurDb,
            double deltaRefTextX = 0.35, double textHight = 0.25, double deltaYref = 0)
        {
            //double deltaRefTextX = 0.35;
            // Open the Block table for read
            BlockTable acBlkTbl;
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
            OpenMode.ForRead) as BlockTable;
         
            // Open the Block table record Model space for write
            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

            // set necessary layer
            MillingDataEngine.Func.Layers.SetAsCurrent("MillingDepth", acCurDb);

            // Create a single-line text object for milling depth
            DBText acText = new DBText();
            acText.SetDatabaseDefaults();
            acText.Position = new Point3d(millingElement.StartPoint.CoordinateX + deltaRefTextX, millingElement.StartPoint.CoordinateY + deltaYref, millingElement.StartPoint.CoordinateZ);
            acText.Height = textHight;
            int roundedMillingDepth = Convert.ToInt32(Math.Round(millingElement.StartMillingDepth, 0));
            if (roundedMillingDepth == 0)
            {
                acText.TextString = roundedMillingDepth + "cm";
            }
            else
            {
                acText.TextString = "-" + roundedMillingDepth + "cm";
            }
            // add text to block table and transaction
            acBlkTblRec.AppendEntity(acText);
            acTrans.AddNewlyCreatedDBObject(acText, true);
        }
    }
}
