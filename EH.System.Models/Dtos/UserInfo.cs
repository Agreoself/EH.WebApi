﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class UserInfo
    {
        public string UserName { get; set; }
        public string FullName { get; set; }
        public int Gender { get; set; }
        public string Department { get; set; }
        public string Email { get; set; }
        public string JobTitle { get; set; }
        public string Report { get; set; }
        public string Phone { get; set; }
        public bool IsActive { get; set; }
        public bool IsAdmin { get; set; }

        public DateTime? StartWorkDate { get; set; }
        public DateTime? EHIStratWorkDate { get; set; }

        public string? CC { get; set; }

        public List<string> RoleList { get; set; }
    }
}
