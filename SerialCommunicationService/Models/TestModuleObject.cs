using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SerialCommunicationService.Models
{
    public class TestModuleObject
    {
        public string serialNumber { get; set; } // value of serialNumber will be GET_SERIAL_NUM command response
        public string inputValues { get; set; } // inputValues will be GET_VALUES and GET_CELL response one by one
        public int id { get; set; }
    }
}
