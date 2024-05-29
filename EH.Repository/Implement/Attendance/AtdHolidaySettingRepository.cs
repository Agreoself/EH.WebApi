using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface.Attendance;

namespace EH.Repository.Implement.Sys
{
    public class ADUserPwdNotifyRepository : RepositoryBase<Atd_HolidaySetting>, IAtdHolidaySettingRepository, ITransient
    {
        public ADUserPwdNotifyRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }
    }
}
