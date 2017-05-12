using System;
using System.Linq;

namespace MillingDataEngine.Func
{
    public class SectionViews
    {
        // find elevation of insert point of the cross section view ---- Is not necessary -----
        public static double ElevationOfLocationPoint(double minSectionViewElevation, double distance)
        {
            int resultToReturn = (int)(minSectionViewElevation - minSectionViewElevation % 1 - distance);
            return resultToReturn;
        }

        //convert station string to station location
        public static double ConvertNameToStationLocation(string nameOfCrossSectionView)
        {
            double station = 0;
            int indexOfPlus = nameOfCrossSectionView.IndexOf("+");
            
            if (indexOfPlus > 1)
            {
                Console.WriteLine();
            }
            int indexOfFreeSpace = nameOfCrossSectionView.IndexOf(" ");
            string stationString = nameOfCrossSectionView.Substring(0, indexOfPlus) + nameOfCrossSectionView.Substring(indexOfPlus + 1, indexOfFreeSpace - 1 - indexOfPlus - 1);
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

    }

}
