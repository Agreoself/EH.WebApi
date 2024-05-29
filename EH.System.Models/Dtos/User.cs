using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class User
    {
        public string? orginUser { get; set; }
        public string username { get; set; }
        public string password { get; set; }
        public bool noPassword { get; set; }

    }
}
