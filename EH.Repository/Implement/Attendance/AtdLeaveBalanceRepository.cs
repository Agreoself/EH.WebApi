using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface.Attendance;

namespace EH.Repository.Implement.Sys
{
    public class AtdLeaveBalanceRepository : RepositoryBase<Atd_LeaveBalance>, IAtdLeaveBalanceRepository, ITransient
    {
        public AtdLeaveBalanceRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }
    }
}
