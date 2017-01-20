using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;

namespace MillingDataEngine.Func
{
    public class Layers
    {
        // add Layer with name and colour
        public static void AddLayers(string name, short colourIndex, Transaction acTrans, LayerTable acLyrTbl)
        {
            if (acLyrTbl.Has(name) == false)
            {
                LayerTableRecord acLyrTblRec = new LayerTableRecord();

                // Assign the layer the ACI color 1 and a name
                acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, colourIndex);
                acLyrTblRec.Name = name;

                // Upgrade the Layer table for write
                acLyrTbl.UpgradeOpen();

                // Append the new layer to the Layer table and the transaction
                acLyrTbl.Add(acLyrTblRec);
                acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
            }
        }

        //set Layer as current Layer
        public static void SetAsCurrent(string layerName, Database acCurDb)
        {
            // Start a transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                OpenMode.ForRead) as LayerTable;
                string sLayerName = layerName;
                if (acLyrTbl.Has(sLayerName) == true)
                {
                    // Set the layer Center current
                    acCurDb.Clayer = acLyrTbl[sLayerName];
                    // Save the changes
                    acTrans.Commit();
                }
            }
        }
    }
}
