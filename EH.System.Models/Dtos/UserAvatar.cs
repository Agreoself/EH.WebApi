using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class UserAvatar
    {
        public Guid ID { get; set; }
        public string Avatar { get; set; }
        public string userId { get; set; }
    }
}
