using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Atd_LeaveSetting:BaseEntity
    {
        /// <summary>
        /// 年假类型
        /// </summary>
        public string LeaveType { get; set; }
        /// <summary>
        /// 资格（男员工、女员工、全部）
        /// </summary>
        public string Qualification { get; set; }

        public decimal? Limit { get; set; }
        public string MinUnit { get; set; }
        public bool IsContainHoliday { get; set; }
        public string? CalculationRule { get; set; }
        public string? Description { get; set; }
        public decimal? NeedVpHour { get; set; }
        public bool NeedHRApprove { get; set; }
        public bool NeedAttachment { get; set; }
    }
}
