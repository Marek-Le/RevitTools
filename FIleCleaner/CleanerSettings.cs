using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.FileCleaner
{
    public class CleanerSettings
    {
        public bool IsActive { get; set; }
        public bool DeleteDirectories { get; set; }
        public int DaysAllowed { get; set; }
        public List<string> DirectoryPaths { get; set; }        
    }
}
