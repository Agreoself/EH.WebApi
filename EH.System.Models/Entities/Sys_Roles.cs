using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Sys_Roles:BaseEntity
    {
        public string RoleName { get; set; }
        public string RoleDescript { get; set; }

        public string RoleKey { get; set;}
        public string? Remark { get; set;}
    }
}
