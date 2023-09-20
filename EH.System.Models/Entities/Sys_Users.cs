using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Sys_Users : BaseEntity
    {
        public string UserName { get; set; }
        public string FullName { get; set; }
        public string? Password { get; set; }
        public int Gender { get; set; }
        public string Department { get; set; }
        public string Email { get; set; }
        public string JobTitle { get; set; }
        public string Report { get; set; }
        public string Phone { get; set; }
        public bool IsActive { get; set; }
        public bool IsAdmin { get; set; }

        public DateTime? StartWorkDate { get; set; }
        public DateTime? EhiStratWorkDate {  get; set; }

        public string? CC { get; set; }

        public override bool Equals(object? obj)
        {
            if (obj is Sys_Users other)
            {
                return UserName == other.UserName && FullName == other.FullName;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(UserName, FullName);
        }

        public class UsersEqulityComparer : IEqualityComparer<Sys_Users>
        {
            public bool Equals(Sys_Users? x, Sys_Users? y)
            {
                return x.UserName == y.UserName && x.FullName == y.FullName;
            }

            public int GetHashCode([DisallowNull] Sys_Users obj)
            {
                return HashCode.Combine(obj.UserName, obj.FullName);
            }
        }

    }
}
