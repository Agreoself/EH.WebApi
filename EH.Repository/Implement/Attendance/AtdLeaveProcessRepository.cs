﻿using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface.Attendance;

namespace EH.Repository.Implement.Sys
{
    public class AtdLeaveProcessRepository : RepositoryBase<Atd_LeaveProcess>, IAtdLeaveProcessRepository, ITransient
    {
        public AtdLeaveProcessRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        } 
    }
}
