using System.Collections.Generic;


namespace MillingDataEngine.DataStruct
{
    public class Cross_section
    {
        public Cross_section(string name, double station, 
            double leftEdgeElev, double rightEdgeElev, double midElev, 
            double width, List<MillingElement> millingElements)
        {
            SlopeLeft = (midElev - leftEdgeElev) / width;
            SlopeRight = (midElev - rightEdgeElev) / width;
            DeltaLevelLeft = midElev - leftEdgeElev;
            DeltaLevelRight = midElev - rightEdgeElev;
            Width = width;
            MillingElements = millingElements;
            Name = name;
        }

        public List<MillingElement> MillingElements { get; private set; }
        public double Width { get; private set; }
        public double SlopeLeft { get; private set; }
        public double SlopeRight { get; private set; }
        public string Name { get; private set; }
        public double Station { get; private set; }
        public double DeltaLevelLeft { get; private set; }
        public double DeltaLevelRight { get; private set; }
    }
}
