using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Sys_UserRole:BaseEntity
    {
        public string UserID { get; set; }
        public string RoleID { get; set; }
    }
}
