using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MillingDataEngine.DataStruct;

namespace MillingDataEngine.DataStruct
{
    public class Cross_section_for_sectionView : Cross_section
    {
        public Cross_section_for_sectionView(Cross_section crossSection, ThreeDPoint insertPointLocation, double insertPointElevation) : 
            base(crossSection.Name, crossSection.Station, crossSection.LeftEdgeElevation, crossSection.RightEdgeElevation, crossSection.MidPoint_Elevation, crossSection.Width, crossSection.MillingElements,
            true, crossSection.MidPointOfCrossSection) 
        {
            base.ProjLayerThick = crossSection.ProjLayerThick;
            this.InsertPoint_Location = insertPointLocation;
            this.InsertPoint_Elevation = insertPointElevation;

            SwapX_and_Y_Coordinates();
            MoveCrossSectionToSectionView();
        }

        public ThreeDPoint InsertPoint_Location { get; private set; }
        public double InsertPoint_Elevation { get; private set; }
        private void MoveCrossSectionToSectionView()
        {
            double deltaX = InsertPoint_Location.CoordinateX - MidPointOfCrossSection.CoordinateX;
            double deltaY = InsertPoint_Location.CoordinateY - MidPointOfCrossSection.CoordinateY;
            double deltaElevation = MidPoint_Elevation - InsertPoint_Elevation;
            if (deltaElevation < 0)
            {
                throw new IndexOutOfRangeException(String.Format("Delta Elevation must be positive value, not {0}!", deltaElevation));
            }
            else
            {
                deltaY = deltaY + deltaElevation;
            }
            MidPointOfCrossSection.CoordinateX = MidPointOfCrossSection.CoordinateX + deltaX;
            MidPointOfCrossSection.CoordinateY = MidPointOfCrossSection.CoordinateY + deltaY;

            foreach (MillingElement singleElement in MillingElements)
            {
                singleElement.StartPoint.CoordinateX = singleElement.StartPoint.CoordinateX + deltaX;
                singleElement.StartPoint.CoordinateY = singleElement.StartPoint.CoordinateY + deltaY;
                singleElement.EndPoint.CoordinateX = singleElement.EndPoint.CoordinateX + deltaX;
                singleElement.EndPoint.CoordinateY = singleElement.EndPoint.CoordinateY + deltaY;
            }
            // We need this because of slope additional change of coordinates of milling elements and other....
            double additionalDeltaY = 0;
            bool test = true;
            foreach (MillingElement singleElement in MillingElements)
            {
                if (singleElement.StartPoint.CoordinateX == MidPointOfCrossSection.CoordinateX)
                {
                    additionalDeltaY = MidPointOfCrossSection.CoordinateY - singleElement.StartPoint.CoordinateY - ProjLayerThick / 100;
                    test = false;
                    break;
                }
                else
                {
                    if (singleElement.EndPoint.CoordinateX == MidPointOfCrossSection.CoordinateX)
                    {
                        additionalDeltaY = MidPointOfCrossSection.CoordinateY - singleElement.EndPoint.CoordinateY - ProjLayerThick / 100;
                        test = false;
                        break;
                    }
                }
            }

            if (test)
            {
                additionalDeltaY = - ProjLayerThick / 100;
            }

            foreach (MillingElement singleElement in MillingElements)
            {
                singleElement.StartPoint.CoordinateY = singleElement.StartPoint.CoordinateY + additionalDeltaY;
                singleElement.EndPoint.CoordinateY = singleElement.EndPoint.CoordinateY + additionalDeltaY;
            }
        }
        private void SwapX_and_Y_Coordinates()
        {
            double swapVariable = MidPointOfCrossSection.CoordinateX;
            MidPointOfCrossSection.CoordinateX = MidPointOfCrossSection.CoordinateY * -1;
            MidPointOfCrossSection.CoordinateY = swapVariable;
            
            foreach (MillingElement singleElement in MillingElements)
            {
                swapVariable = singleElement.StartPoint.CoordinateX;
                singleElement.StartPoint.CoordinateX = singleElement.StartPoint.CoordinateY * -1;
                singleElement.StartPoint.CoordinateY = swapVariable;

                swapVariable = singleElement.EndPoint.CoordinateX;
                singleElement.EndPoint.CoordinateX = singleElement.EndPoint.CoordinateY * -1;
                singleElement.EndPoint.CoordinateY = swapVariable;
            }
        }
    }
}
