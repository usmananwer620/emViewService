using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SerialCommunicationService
{
    class ModuleResponseResult
    {
        public long SitesId { get; set; }
        public string SerialNumber { get; set; }
        public string Name { get; set; }
        public string GET_FW_DATE { get; set; }
        public string GET_TND { get; set; }
    }
}
