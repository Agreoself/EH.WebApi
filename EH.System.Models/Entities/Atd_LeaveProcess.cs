using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Atd_LeaveProcess:BaseEntity
    {
     
        public string ProcessId { get; set; }
 
        public string UserId { get; set; }

        public string LeaveId { get; set; }
        public string? Action { get; set; }
        public string? Result { get; set; }
        public DateTime? AuditTime { get; set; }
        public int OrderNo { get; set; }
        public string? ProcessState { get; set; } 
        public bool IsLastNode { get; set; }

    }
}
