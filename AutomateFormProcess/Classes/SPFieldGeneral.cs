using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomateFormProcess.Classes
{
    public class SPFieldGeneral
    {
        // Properties
        public string AggregationFunction { get; set; }

        public bool AllowMultipleValue { get; set; }

        public string DefaultValue { get; set; }

        public string Description { get; set; }

        public Guid Guid { get; set; }

        public string InternalName { get; set; }

        public bool IsFile { get; set; }

        public bool IsRequire { get; set; }

        public string LookupList { get; set; }

        public string LookupTitleField { get; set; }

        public string LookupValueField { get; set; }

        public string MasterLookupName { get; set; }

        public int MaxLength { get; set; }

        public double MaxValue { get; set; }

        public double MinValue { get; set; }

        public List<string> options { get; set; }

        public string Query { get; set; }

        public string[] RelatedFields { get; set; }
        public string NotShow { get; set; }
        public string Disable { get; set; }

        public bool ShowAsPercentage { get; set; }

        public string Title { get; set; }

        public string Type { get; set; }

        public string TypeFile { get; set; }

        public string VolumeFile { get; set; }
        public string Status { get; set; }
    }
}
