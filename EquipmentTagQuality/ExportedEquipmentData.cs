using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.EquipmentTagQuality
{
    public class ExportedEquipmentData
    {
        public string revit_element_id { get; set; }
        public string revit_element_guid { get; set; }
        public string equipment_tag { get; set; }
        public string bounding_box { get; set; }
        public string element_data { get; set; }
        public string model_file_name { get; set; }
        public string model_file_guid { get; set; }
        public string nearest_grid_intersection { get; set; }
    }
}
