using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Repository.Implement.Sys
{
    public class SysUsersRepository : RepositoryBase<Sys_Users>,ISysUsersRepository, ITransient
    {
        public SysUsersRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }

        public bool isAdmin(Guid userId)
        {
            return base.FirstOrDefault(i => i.ID == userId).IsAdmin;
        }
    }
}
