using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MillingDataEngine.DataStruct
{
    public class MillingElement
    {
        // in feature if necessary to make different scales in axes X or Y or Z
        private double scaleX = 1;
        private double scaleY = 1;
        private double scaleZ = 1; // not used for now

        // represent milling Ranges
        private static double [] millingRange_1 = {0d, 3d};
        private static double[] millingRange_2 = {3d, 5d};
        private static double [] millingRange_3 = {5d, 7d};
        private static double[] millingRange_h7 = { 7d, -1d };

        private double[] rangeScope;

        private double refStart = -50;


        // posible name of layers
        private string[] layerNames = { "StA_Strugane_0_3", "StA_Strugane_3_5", "StA_Strugane_5_7", "StA_Strugane_7" };
        
        public MillingElement(double station, string profileName, double lineStart, double lineLength, double startMillingDepth, double endMillingDepth)
        {
            if (profileName == "206")
            {
                Console.WriteLine();
            }
            Station = station;
            LineStart = lineStart;
            LineLength = lineLength;
            StartMillingDepth = startMillingDepth;
            EndMillingDepth = endMillingDepth;
            ProfileName = profileName;

            SetLayerName(startMillingDepth, endMillingDepth);
            SetStartPoint();
            SetEndPoint();
        }

        public double Station { get; set; }
        public string ProfileName { get; set; }
        public double LineStart { get; set; }
        public double LineLength { get; set; }
        public double StartMillingDepth { get; set; }
        public double EndMillingDepth { get; set; }
        public string LayerName { get; private set; }
        public ThreeDPoint StartPoint { get; private set; }
        public ThreeDPoint EndPoint { get; private set; }
        public static double[] MillingRange_1 { get { return millingRange_1; } }
        public static double[] MillingRange_2 { get { return millingRange_2; } }
        public static double[] MillingRange_3 { get { return millingRange_3; } }
        public static double[][] MillingRanges { get { return new double[][] { millingRange_1, millingRange_2, millingRange_3 }; } } //return all milling ranges
        public double[] RangeScope { get { return rangeScope; } }
        public double RefStart { get { return refStart; } }

        //set layer name
        private void SetLayerName(double startMilling, double endMilling)
        {
            double millingAverage = Math.Abs(startMilling + endMilling) / 2;
            if (millingAverage >= millingRange_1[0] && millingAverage < millingRange_1[1]) // see if first statement is bether with '>='
            {
                LayerName = layerNames[0];
                rangeScope = MillingRange_1;
            }
            else
            {
                if (millingAverage >= millingRange_2[0] && millingAverage < millingRange_2[1])
                {
                    LayerName = layerNames[1];
                    rangeScope = MillingRange_2;                    
                }
                else
                {
                    if (millingAverage >= millingRange_3[0] && millingAverage < millingRange_3[1])
                    {
                        LayerName = layerNames[2];
                        rangeScope = MillingRange_3;
                        
                    }
                    else
                    {
                        if (millingAverage >= millingRange_3[1])
                        {
                            LayerName = layerNames[3];
                            rangeScope = millingRange_h7;
                        }
                        else
                        {
                           throw new System.ArgumentOutOfRangeException("Milling can not be negative value!");
                        }
                    }
                }
            }
        }

        // set start point
        private void SetStartPoint()
        {
            StartPoint = new ThreeDPoint((Station + refStart) * scaleX, (LineStart - refStart) * scaleY);
        }

        // set end point
        private void SetEndPoint()
        {
            EndPoint = new ThreeDPoint((Station + refStart) * scaleX, (LineStart - LineLength - refStart) * scaleY);
        }
    }
}
