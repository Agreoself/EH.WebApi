using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class UserRole
    {
        public List<string> userIds { get; set; }
        public List<string> roleIds { get; set; }
    }
}
