using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.EquipmentTagQuality.GigabaseModels
{
    public class SystemType
    {
        public int id { get; set; }
        public string system_name { get; set; }
        public string system_abbreviation { get; set; }
        public string system_code { get; set; }
        public string status { get; set; }
        public int discipline_id { get; set; }
        public string discipline_abbreviation { get; set; }
        public string discipline_name { get; set; }
    }
}
