using Autodesk.Revit.DB;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TeslaRevitTools.ITwoExport
{
    public class ExcelMinWrapper
    {
        public static DataSet GetExcelDataSet(string excelFilePath)
        {
            DataSet result = null;
            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read())
                        {
                            //p
                        }
                    } while (reader.NextResult());
                    result = reader.AsDataSet();
                }
            }
            return result;
        }
    }
}
