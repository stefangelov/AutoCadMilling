using System;
using System.Linq;

namespace MillingDataEngine.Func
{
    public class SectionViews
    {
        // find elevation of insert point of the cross section view
        public static double ElevationOfLocationPoint(double minSectionViewElevation, double distance)
        {
            return (int)(minSectionViewElevation - minSectionViewElevation % 10);
        }

        //convert station string to station location
        public static double ConvertNameToStationLocation(string nameOfCrossSectionView)
        {
            double station = 0;
            int indexOfPlus = nameOfCrossSectionView.IndexOf("+");
            int indexOfFreeSpace = nameOfCrossSectionView.IndexOf(" ");
            string stationString = nameOfCrossSectionView.Substring(0, indexOfPlus) + nameOfCrossSectionView.Substring(indexOfPlus + 1, indexOfFreeSpace-1);
            try
            {
                station = Convert.ToDouble(stationString);
            }
            catch (Exception)
            {
                throw new ArgumentException("Value of Station is not correct!");
            }

            return station;
        }

        public static bool IsSTationsMatch (MillingDataEngine.DataStruct.Cross_section crossSection, double searchStation, double range = 0.50d)
        {
            return (crossSection.MillingElements.Count > 0 &&
                ((crossSection.MillingElements[0].Station >= searchStation - range || crossSection.MillingElements[0].Station >= searchStation + range)));
        }

    }

}
