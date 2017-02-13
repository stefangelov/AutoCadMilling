using System;
namespace MillingDataEngine.Func
{
    public class SectionViews
    {
        public static double ElevationOfLocationPoint(double minSectionViewElevation)
        {
            double elevationOfLocationPoint = -1;

            elevationOfLocationPoint = (int) (minSectionViewElevation - minSectionViewElevation % 10 - 10);

            return elevationOfLocationPoint;
        }
    }
}
