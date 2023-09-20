using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Repository.Implement.Sys
{
    public class SysRolesRepository : RepositoryBase<Sys_Roles>,ISysRolesRepository, ITransient
    {
        public SysRolesRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }
    }
}
