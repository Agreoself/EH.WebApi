using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface.Attendance;

namespace EH.Repository.Implement.Sys
{
    public class AtdLeaveBalanceBakRepository : RepositoryBase<Atd_LeaveBalance_Bak>, IAtdLeaveBalanceBakRepository, ITransient
    {
        public AtdLeaveBalanceBakRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }
    }
}
