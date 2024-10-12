using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Atd_LeaveBalance:BaseEntity
    {
        
        public string UserId { get; set; }
      
        public string LeaveType { get; set; }
        public int Year { get; set; }

        public decimal Total { get; set; }
        public decimal Available { get; set; }
        public decimal Used { get; set; }
        public decimal Locked { get; set; }
        public decimal AVCarryoverTotal { get; set; }
        public decimal AVCarryoverAvailable { get; set; }
        public decimal AVCarryoverUsed { get; set; }
        public decimal AVCarryoverLocked { get; set; }
        public decimal AVCarryoverCleared { get; set; }
        public bool IsClear { get; set; }
        public string? Remark { get; set; }
   
    }
}
