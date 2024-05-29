using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Atd_ApproveByEmail
    {
        public string MoveEmail { get; set; }
        public string MessageId { get; set; }
        public string FromEmail { get; set; }
        public string LeaveId { get; set; }
        public string Result { get; set; }
        public string Comment { get; set; }

    }

    public class MoveInfo
    {
        public string MessageID { get; set; }
        public string Email { get; set; }
    }
}
