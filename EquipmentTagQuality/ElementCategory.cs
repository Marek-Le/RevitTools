using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public class ElementCategory
    {
        public bool IsChecked { get; set; }
        public string ElementCategoryName { get; set; }
        public BuiltInCategory BuiltInCategory { get; set; }
    }
}
