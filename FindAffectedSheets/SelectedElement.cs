using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.FindAffectedSheets
{
    public class SelectedElement
    {
        public int ElementId { get; set; }
        public string TypeName { get; set; }
        public string CategoryName { get; set; }
        public Element Element { get; set; }
    }
}
