using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MWO_XMLReader
{
    static class ConfigFile
    {
        public static bool Initialize()
        {

            return true;
        }

        public static int LOG_LEVEL {get;set;} = 6;
        public static string DIR_QUIRK { get; set; } = "C:\\DOCUMENTI\\MWO\\Quirk";
        public static string QUIRK_PAGE { get; set; } = "IS Quirk";
        public static string OUTPUT_PAGE { get; set; } = "MechLab";
        public static FileInfo TEMPLATE = new FileInfo("C:\\TEST\\MWO worksheet.xlsx");
    }
}
