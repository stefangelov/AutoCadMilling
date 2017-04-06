using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MillingDataEngine.DataStruct
{
    public class MIllingQuantity
    {
        public MIllingQuantity(double[] range, double length, string layer)
        {
            Range = range;
            if (length < 0)
            {
                throw new ArgumentOutOfRangeException("Length must be positive value!");
            }
            else
            {
                Quant = length;
            }
            LayerName = layer;
        }
        public double [] Range { get; set; }
        public double Quant { get; set; }
        public string LayerName { get; private set; }
    }
}
