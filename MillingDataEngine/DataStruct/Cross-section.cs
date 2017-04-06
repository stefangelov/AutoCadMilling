using System.Collections.Generic;
using MillingDataEngine.DataStruct;
using System.Linq;
using System;


namespace MillingDataEngine.DataStruct
{
    public class Cross_section
    {
        public Cross_section(string name, double station, 
            double leftEdgeElev, double rightEdgeElev, double midElev, 
            double width, List<MillingElement> millingElements, bool isToSectionView = false, ThreeDPoint midPointOfCrossSection = null)
        {
            DeltaLevelLeft = midElev - leftEdgeElev;
            DeltaLevelRight = midElev - rightEdgeElev;
            SlopeLeft = (DeltaLevelLeft) / (width / 2);
            SlopeRight = (DeltaLevelRight) / (width / 2);
            Width = width;
            MillingElements = millingElements;
            Name = name;
            LeftEdgeElevation = leftEdgeElev;
            RightEdgeElevation = rightEdgeElev;

            MidPoint_Elevation = midElev;
            
            if (!isToSectionView)
            {
                ChangeStartEndPointForTheDiagram();
            }
            if (midPointOfCrossSection != null)
            {
                MidPointOfCrossSection = midPointOfCrossSection;
            }
        }


        public double ProjLayerThick { get; set; }
        public double LeftEdgeElevation { get; private set; }
        public double RightEdgeElevation { get; private set; }
        public List<MillingElement> MillingElements { get; private set; }
        public double Width { get; private set; }
        public double SlopeLeft { get; private set; }
        public double SlopeRight { get; private set; }
        public string Name { get; private set; }
        public double Station { get; private set; }
        public double DeltaLevelLeft { get; private set; }
        public double DeltaLevelRight { get; private set; }
        public List<MIllingQuantity> MillingQuantity { get { return CalcMillingQuantities(); } }
        public ThreeDPoint MidPointOfCrossSection { get; private set; }
        public double MidPoint_Elevation { get; private set; }

        private List<MIllingQuantity> CalcMillingQuantities()
        {
            List<MIllingQuantity> quantityForReturning = new List<MIllingQuantity>();
           
            int countOfAllMillingElements = MillingElements.Count;
            quantityForReturning.Add(new MIllingQuantity(MillingElements[0].RangeScope, MillingElements[0].LineLength, MillingElements[0].LayerName));

            for (int i = 1; i < countOfAllMillingElements; i++)
            {
                if (quantityForReturning.Exists(x => x.Range == MillingElements[i].RangeScope))
                {
                    quantityForReturning.Find(x => x.Range == MillingElements[i].RangeScope).Quant += MillingElements[i].LineLength;
                }
                else
                {
                    quantityForReturning.Add(new MIllingQuantity(MillingElements[i].RangeScope, MillingElements[i].LineLength, MillingElements[i].LayerName));
                }
            }
            List<MIllingQuantity> sortedQuantList = quantityForReturning.OrderByDescending(x => x.Range[0]).ToList();
            return sortedQuantList;
        }

        private void ChangeStartEndPointForTheDiagram()
        { 
            // first change Y cordinate according to center line
            foreach (MillingElement singleElement in MillingElements)
            {
                // change the Y coordinate
                singleElement.StartPoint.CoordinateY = singleElement.StartPoint.CoordinateY + Width / 2;
                singleElement.EndPoint.CoordinateY = singleElement.EndPoint.CoordinateY + Width / 2;
            }
            // find mid point
            MidPointOfCrossSection = FindMidPoint();
            ThreeDPoint midPointOfCrossSection = MidPointOfCrossSection;
            // change X coordinate of each start and end point to correspond to slope
            foreach (MillingElement singleElement in MillingElements)
            {
                // change the X coordinate
                singleElement.StartPoint.CoordinateX = singleElement.StartPoint.CoordinateX - FindDeltaAcordingToSlope(midPointOfCrossSection, singleElement.StartPoint);
                singleElement.EndPoint.CoordinateX = singleElement.EndPoint.CoordinateX - FindDeltaAcordingToSlope(midPointOfCrossSection, singleElement.EndPoint);      
            }
        }

        private double FindDeltaAcordingToSlope(ThreeDPoint midPoint, ThreeDPoint thePoint)
        {
            double delta = 0;

            if (midPoint.CoordinateY != thePoint.CoordinateY)
            {
                if (midPoint.CoordinateY < thePoint.CoordinateY)
                {
                    delta = (thePoint.CoordinateY - midPoint.CoordinateY) * SlopeLeft;
                }
                else
                {
                    delta = (midPoint.CoordinateY - thePoint.CoordinateY) * SlopeRight;
                }
            }

            return delta;
        }

        // find mid point of cross section
        private ThreeDPoint FindMidPoint()
        {
            double midX;
            double midY;
            if (MillingElements.Count > 0)
            {
                midY = Math.Abs(MillingElements[0].RefStart);
                //midY = (MillingElements[0].StartPoint.CoordinateY + MillingElements[MillingElements.Count - 1].EndPoint.CoordinateY) / 2;
                midX = MillingElements[0].StartPoint.CoordinateX;
                ThreeDPoint midPoint = new ThreeDPoint(midX, midY);
                return midPoint;
            }
            else
            {
                return null;
            }
        }
    }
}
