using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MWO_XMLReader
{
    public class MechStats
    {
        public string Variant { get; set; } = string.Empty;
        public string Chassis { get; set; } = string.Empty;
        public int Class { get; set; } = 0;
        public int MaxTons { get; set; } = 0;
        public int MaxJumpJets { get; set; } = 0;
        public bool CanEquipECM { get; set; } = false;
        public int MinEngineRating { get; set; } = 0;
        public int MaxEngineRating { get; set; } = 0;
        public List<Quirk> QuirkList { get; set; } = new List<Quirk>();
            

    }
}
