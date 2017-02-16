using System.Collections.Generic;


namespace MillingDataEngine.DataStruct
{
    public class Cross_section
    {
        public Cross_section(string name, double station, 
            double leftEdgeElev, double rightEdgeElev, double midElev, 
            double width, List<MillingElement> millingElements)
        {
            SlopeLeft = (midElev - leftEdgeElev) / width / 2;
            SlopeRight = (midElev - rightEdgeElev) / width / 2;
            DeltaLevelLeft = midElev - leftEdgeElev;
            DeltaLevelRight = midElev - rightEdgeElev;
            Width = width;
            MillingElements = millingElements;
            Name = name;

            ChangeStartEndPointForTheDiagram();
        }
        public List<MillingElement> MillingElements { get; private set; }
        public double Width { get; private set; }
        public double SlopeLeft { get; private set; }
        public double SlopeRight { get; private set; }
        public string Name { get; private set; }
        public double Station { get; private set; }
        public double DeltaLevelLeft { get; private set; }
        public double DeltaLevelRight { get; private set; }

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
            ThreeDPoint midPointOfCrossSection = FindMidPoint();
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
            if (MillingElements.Count > 0)
            {
                double midX = MillingElements[0].StartPoint.CoordinateX;
                double midY = (MillingElements[0].StartPoint.CoordinateY + MillingElements[MillingElements.Count - 1].EndPoint.CoordinateY) / 2;
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
