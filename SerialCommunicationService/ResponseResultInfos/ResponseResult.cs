using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SerialCommunicationService
{
    class ResponseResult
    {
        public string accessToken { get; set; }
        public string encryptedAccessToken { get; set; }
        public string expireInSeconds { get; set; }
        public string userId { get; set; }
    }
}
