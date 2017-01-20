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
        public List<Cross_section> CrossSections { get { return theCrossSections; } }

        public void AddCross(Cross_section crossSection)
        {
            theCrossSections.Add(crossSection);
        }
    }
}
