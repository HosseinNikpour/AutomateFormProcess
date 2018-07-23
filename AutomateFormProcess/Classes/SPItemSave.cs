using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace AutomateFormProcess.Classes
{
    public class SPItemSave
    {
        public string ListId { get; set; }
        public SPListItem Item { get; set; }
        public SPFolder Folder { get; set; }
    }
}
