using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Atd_LeaveBalance_Bak : BaseEntity
    {
        
        public string UserId { get; set; }
      
        public string LeaveType { get; set; }
        public int Year { get; set; }

        public decimal Total { get; set; }
        public decimal Available { get; set; }
        public decimal Used { get; set; }
        public decimal Locked { get; set; }
        public decimal? Overdue { get; set; }
        public string? Remark { get; set; }
        public decimal AnnualClear { get; set; }

    }
}
