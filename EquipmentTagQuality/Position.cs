using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public class Position
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

        public Position()
        {

        }

        public Position(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }
    }
}
