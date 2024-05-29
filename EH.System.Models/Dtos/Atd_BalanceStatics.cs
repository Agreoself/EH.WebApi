using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Atd_BalanceStatics
    {
        public string UserName { get; set; }
        public string ChineseName { get; set; }
        public string Department { get; set; }
        public decimal? AnnualTotal { get; set; }
        public decimal? AnnualAvailable { get; set; }
        public decimal? AnnualCarryover { get; set; }
        public decimal? AnnualCarryoverUsed { get; set; }
        
        public decimal? AnnualUsed { get; set; }
        public decimal? AnnualLocked { get; set; }

        public decimal? Sick { get; set; }
        public decimal? Personal { get; set; }
        public decimal? Wedding { get; set; }
        public decimal? Paternity { get; set; }
        public decimal? Parental { get; set; }

        public decimal? Prenatal { get; set; }
        public decimal? Abortion { get; set; }
        public decimal? Bereavement { get; set; }
        public decimal? Maternity { get; set; }

        public decimal? Nursing { get; set; }
        public decimal? Marriage { get; set; }
        public decimal? PrenatalCheckup { get; set; }


    }
}
