using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Atd_Audit
    {
         
        public string ProcessID { get; set; }
        public string FormID { get; set; }
        public string Result { get; set; }
        public string? Comment { get; set; }

    }
}
