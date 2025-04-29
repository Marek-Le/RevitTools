using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public class RevitTransform
    {
        public Position BasisX { get; set; }
        public Position BasisY { get; set; }
        public Position BasisZ { get; set; }
        public Position Origin { get; set; }

        public void CreateDefaultTransform()
        {
            BasisX = new Position(1, 0, 0);
            BasisY = new Position(0, 1, 0);
            BasisZ = new Position(0, 0, 1);
            Origin = new Position(0, 0, 0);
        }

        public void CreateCustomTransform(Position origin, Position basisX, Position basisY, Position basisZ)
        {
            BasisX = basisX;
            BasisY = basisY;
            BasisZ = basisZ;
            Origin = origin;
        }

        public Position TransformRelativePosition(Position position)
        {
            double x = BasisX.X * position.X + BasisY.X * position.Y + BasisZ.X * position.Z;
            double y = BasisX.Y * position.X + BasisY.Y * position.Y + BasisZ.Y * position.Z;
            double z = BasisX.Z * position.X + BasisY.Z * position.Y + BasisZ.Z * position.Z;
            return new Position(x, y, z);
        }

        public Position TransformAbsolutePosition(Position position)
        {
            double x = BasisX.X * (position.X + Origin.X) + BasisY.X * (position.Y + Origin.Y) + BasisZ.X * (position.Z + Origin.Z);
            double y = BasisX.Y * (position.X + Origin.X) + BasisY.Y * (position.Y + Origin.Y) + BasisZ.Y * (position.Z + Origin.Z);
            double z = BasisX.Z * (position.X + Origin.X) + BasisY.Z * (position.Y + Origin.Y) + BasisZ.Z * (position.Z + Origin.Z);
            return new Position(x, y, z);
        }
    }
}
