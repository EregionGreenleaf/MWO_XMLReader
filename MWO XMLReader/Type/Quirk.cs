using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MWO_XMLReader
{
    public class Quirk
    {
        public string Name { get; set; } = string.Empty;
        public bool State { get; set; } = false;
        public double Value { get; set; } = 0.0;
    }
}
