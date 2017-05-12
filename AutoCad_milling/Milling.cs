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
using Autodesk.AECC.Interop.Roadway;
using System.IO;
using System.Linq;

using Section = Autodesk.Civil.DatabaseServices.Section;


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
                MillingDataEngine.Func.Layers.AddLayers("Axis", 1, acTrans, acLyrTbl); // After drawing creation You must change line type manual (for now).
                MillingDataEngine.Func.Layers.AddLayers("Texts", 7, acTrans, acLyrTbl);
                MillingDataEngine.Func.Layers.AddLayers("MillingDepth", 6, acTrans, acLyrTbl);

                
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;

                // create milling Elements
                MillingDataEngine.DataStruct.RoadSection theRoadSection = MillingDataEngine.Func.ExcelDataRead.RoadSectionElementsBuilder(MillingDataEngine.Func.ExcelDataRead.ReadData());
                
                // drow Elements
                MillingDataEngine.Func.DrawMillingElements.Drow(theRoadSection, acTrans, acBlkTblRec, acCurDb);
                
                // Save the changes and dispose of the transaction
                acTrans.Commit();
            }
        }


         // the name of Acad Command to insert milling depth on section profile views
        [CommandMethod("MillingDepthToProfileView")]
        [STAThread]
        public static void MillingDepthToProfileView()
        {
            CivilDocument civilDoc = CivilApplication.ActiveDocument;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            List<string> resultForFile = new List<string>();

            using (Transaction ts = Application.DocumentManager.MdiActiveDocument.
                Database.TransactionManager.StartTransaction())
                {
                // dterminate location of section view profiles
                // ask user to select Alignment
                PromptEntityOptions opt = new PromptEntityOptions("\nSelect an Alignment");
                opt.SetRejectMessage("\nObject must be an alignment.\n");
                opt.AddAllowedClass(typeof(Alignment), false);
                ObjectId alignID = ed.GetEntity(opt).ObjectId;
               
                try
                {
                    Alignment myAlignment = ts.GetObject(alignID, OpenMode.ForRead) as Alignment;
                    ObjectIdCollection sampleLineIdCollection = myAlignment.GetSampleLineGroupIds();
                    SampleLineGroup sampleLineGroup = sampleLineIdCollection[0].GetObject(OpenMode.ForRead) as SampleLineGroup; //if you have problem chech here

                    foreach (ObjectId sampleLineId in sampleLineGroup.GetSampleLineIds())
                    {
                        SampleLine sampleLine = ts.GetObject(sampleLineId, OpenMode.ForRead) as SampleLine;

                        foreach (ObjectId sectionId in sampleLine.GetSectionIds())
                        {
                            Section section = ts.GetObject(sectionId, OpenMode.ForWrite) as Section;
                            ed.WriteMessage("Section {0} elevation max: {1} min: {2}\n", section.Name, section.MaximumElevation, section.MinmumElevation);
                            // set the section update mode:
                            //section.UpdateMode = SectionUpdateType.Dynamic;
                        }
                    }

                    SectionViewGroupCollection sectionViewGrouopColection = sampleLineGroup.SectionViewGroups;

                    int sectionViewNumber = 0;
                    try
                    {
                        string askString = String.Format("\nEnter number of ACTIVE SectionViewGroup (1-{0}): ", sectionViewGrouopColection.Count);
                        sectionViewNumber = Convert.ToInt32(GetStringFromUser(ed, askString));
                        if (sectionViewNumber < 1 || sectionViewNumber > sectionViewGrouopColection.Count)
                        {
                            throw new ArgumentOutOfRangeException("Value is out of range!");
                        }
                    }
                    catch (System.Exception)
                    {
                        ed.WriteMessage("Please enter valid value!");
                        return;
                    }
                    SectionViewGroup theSectionViewGroup = sectionViewGrouopColection[sectionViewNumber - 1];
                    ObjectIdCollection sectionViewIdColection = theSectionViewGroup.GetSectionViewIds();

                    // create milling Elements
                    MillingDataEngine.DataStruct.RoadSection theRoadSection = MillingDataEngine.Func.ExcelDataRead.RoadSectionElementsBuilder(MillingDataEngine.Func.ExcelDataRead.ReadData());

                    foreach (ObjectId sectionViewId in sectionViewIdColection)
                    {
                        SectionView theSectionView = ts.GetObject(sectionViewId, OpenMode.ForWrite) as SectionView;

                        theSectionView.IsElevationRangeAutomatic = false;
                        string minRangeElevation = theSectionView.ElevationMin.ToString();
                        theSectionView.IsElevationRangeAutomatic = true;

                        MillingDataEngine.DataStruct.ThreeDPoint theLocation = new MillingDataEngine.DataStruct.ThreeDPoint(theSectionView.Location.X, theSectionView.Location.Y, Convert.ToInt32(theSectionView.Location.Z));
                        double baseElevation = Convert.ToDouble(minRangeElevation);

                        double station = MillingDataEngine.Func.SectionViews.ConvertNameToStationLocation(theSectionView.Name);

                        double range = 0.5; //+- range of stations

                        // search maching of cross section in drowing ana xls tanble
                        MillingDataEngine.DataStruct.Cross_section tempCross = theRoadSection.CrossSections.Find((x) => ((x.Station >= station - range && x.Station <= station) ||
                                                                                                                        (x.Station <= station + range && x.Station >= station)));
                        if (tempCross != null && tempCross.MillingElements.Count > 0) //if there is milling elements go next
                        {
                            MillingDataEngine.DataStruct.Cross_section_for_sectionView tempCrossForCV = new MillingDataEngine.DataStruct.Cross_section_for_sectionView(tempCross, theLocation, baseElevation);
                            theRoadSection.AddCrossSectionview(tempCrossForCV);
                        } 
                    }

                    DrowToSectionViews(theRoadSection);

                }
                catch (System.Exception e)
                {
                    //throw new System.Exception(e.Message);
                    ed.WriteMessage(e.Message);
                    //return;
                }
                  
                // Save the changes and dispose of the transaction
                ts.Commit();
            }
            ed.WriteMessage("\nOK!");
        }

        private static void DrowToSectionViews(MillingDataEngine.DataStruct.RoadSection theRoadSection)
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
                MillingDataEngine.Func.Layers.AddLayers("Axis", 1, acTrans, acLyrTbl); // After drawing creation You must change line type manual (for now).
                MillingDataEngine.Func.Layers.AddLayers("Texts", 7, acTrans, acLyrTbl);
                MillingDataEngine.Func.Layers.AddLayers("MillingDepth", 6, acTrans, acLyrTbl);


                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;

                // drow Elements
                MillingDataEngine.Func.DrawMillingElements.Drow(theRoadSection, acTrans, acBlkTblRec, acCurDb, true);

                // Save the changes and dispose of the transaction
                acTrans.Commit();
            }
        }

        // Ask User to Input some data
        public static string GetStringFromUser(Editor ed, string askingString)
        {
            PromptStringOptions pStrOpts = new PromptStringOptions(askingString);
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = ed.GetString(pStrOpts);

            return pStrRes.StringResult;
        }
    }
}
