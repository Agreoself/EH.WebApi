using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class LeaveSettingRequest
    {
        public string userId { get; set; }
        public string leaveType { get; set; }
    }
}
