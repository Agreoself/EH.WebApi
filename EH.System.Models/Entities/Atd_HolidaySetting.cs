using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Atd_HolidaySetting:BaseEntity
    {
        public DateTime HolidayDate { get; set; }
        public string? Remark { get; set; }
    }
}
