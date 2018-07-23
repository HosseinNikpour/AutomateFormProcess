using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomateFormProcess.Classes
{
    public class SPFieldValue
    {
        // Properties
        public string InternalName { get; set; }

        public string LookupList { get; set; }

        public List<List<SPFieldValue>> rows { get; set; }

        public string Type { get; set; }

        public string value { get; set; }
    }

 
}
