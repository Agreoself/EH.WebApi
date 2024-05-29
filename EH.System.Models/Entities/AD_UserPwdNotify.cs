using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class AD_UserPwdNotify : BaseEntity
    {
        public string UserId { get; set; }
        public string? Department { get; set; }
        public string? Location { get; set; }
        public string? Email { get; set; }
        public DateTime? PwdLastSetDate { get; set; }
        public DateTime? ExpiredDate { get; set; }
        public string? Remark { get; set; }

        public override bool Equals(object? obj)
        {
            if (obj is AD_UserPwdNotify other)
            {
                return UserId == other.UserId && Email == other.Email;
            }
            return false;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(UserId, Email);
        }

        public class UsersEqulityComparer : IEqualityComparer<AD_UserPwdNotify>
        {
            public bool Equals(AD_UserPwdNotify? x, AD_UserPwdNotify? y)
            {
                return x.UserId == y.UserId && x.Email == y.Email&&x.ExpiredDate==y.ExpiredDate;
            }

            public int GetHashCode([DisallowNull] AD_UserPwdNotify obj)
            {
                return HashCode.Combine(obj.UserId, obj.Email,obj.ExpiredDate);
            }
        }
    }
}
