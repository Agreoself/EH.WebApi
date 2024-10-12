using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{

    public class LeaveDetail
    {
        public int Code { get; set; }
        public string LeaveType { get; set; }
        public decimal? Available { get; set; }

        public string MinUnit { get; set; }

        public bool IsHoliday { get; set; }
        public bool needHr { get; set; }

        public Dictionary<int, List<string>>  holidays { get; set; }
        public bool needAttachment { get; set; }
    }
}
