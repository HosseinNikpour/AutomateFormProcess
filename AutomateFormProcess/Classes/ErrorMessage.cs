using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomateFormProcess.Classes
{
    public class ErrorMessage
    {
        // Properties
        public List<string> FieldNames { get; set; }

        public string Message { get; set; }

        public int RowNumber { get; set; }
    }
}
