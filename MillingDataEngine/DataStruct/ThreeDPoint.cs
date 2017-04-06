using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MillingDataEngine.DataStruct
{
    public class ThreeDPoint
    {
        public ThreeDPoint(double x, double y, int z = 0)
        {
            CoordinateX = x;
            CoordinateY = y;
            CoordinateZ = z;
        }
        public double CoordinateX { get; set; }
        public double CoordinateY { get; set; }
        public double CoordinateZ { get; set; }
    }
}
