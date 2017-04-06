using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MillingDataEngine.DataStruct
{
    public class RoadSection
    {
        private List<MillingDataEngine.DataStruct.Cross_section> theCrossSections = new List<MillingDataEngine.DataStruct.Cross_section>();
        private List<MillingDataEngine.DataStruct.Cross_section_for_sectionView> theCrossSectionsForSectionViews = new List<MillingDataEngine.DataStruct.Cross_section_for_sectionView>();
        public List<Cross_section> CrossSections { get { return theCrossSections; } }

        public List<Cross_section_for_sectionView> CrossSectionsForSectionView { get { return theCrossSectionsForSectionViews; } }

        public void AddCross(Cross_section crossSection)
        {
            theCrossSections.Add(crossSection);
        }

        public void AddCrossSectionview(Cross_section_for_sectionView crossSection)
        {
            theCrossSectionsForSectionViews.Add(crossSection);
        }
    }
}
