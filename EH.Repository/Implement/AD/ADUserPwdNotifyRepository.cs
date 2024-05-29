using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface.Attendance;
using EH.Repository.Implement.Sys;
using EH.Repository.Interface.AD;

namespace EH.Repository.Implement.AD
{
    public class ADUserPwdNotifyRepository : RepositoryBase<AD_UserPwdNotify>, IADUserPwdNotifyRepository, ITransient
    {
        public ADUserPwdNotifyRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }
    }
}
