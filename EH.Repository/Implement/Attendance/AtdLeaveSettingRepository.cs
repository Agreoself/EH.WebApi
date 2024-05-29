using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface.Attendance;
using System.Security.Cryptography.X509Certificates;
using EH.System.Models.Dtos;

namespace EH.Repository.Implement.Sys
{
    public class AtdLeaveSettingRepository : RepositoryBase<Atd_LeaveSetting>, IAtdLeaveSettingRepository, ITransient
    {
       

        public AtdLeaveSettingRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {
         
        }

     
        
    }
}
