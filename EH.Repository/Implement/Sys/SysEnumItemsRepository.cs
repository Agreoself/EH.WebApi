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
    public class SysEnumItemsRepository : RepositoryBase<Sys_EnumItem>,ISysEnumItemsRepository, ITransient
    {
        public SysEnumItemsRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }
    }
}
