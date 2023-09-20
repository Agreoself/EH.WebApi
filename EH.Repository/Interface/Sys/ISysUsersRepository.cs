﻿using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Repository.Interface.Sys
{
    public interface ISysUsersRepository:IRepositoryBase<Sys_Users>
    { 
        public bool isAdmin(Guid userId);
    }
}
