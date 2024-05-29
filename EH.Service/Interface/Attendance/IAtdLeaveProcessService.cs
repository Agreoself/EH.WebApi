﻿using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.Attendance
{
    public interface IAtdLeaveProcessService : IBaseService<Atd_LeaveProcess>
    {
        bool AddProcessByLeaveType(string leaveId);

        void GetUnapprovedUser();
    }
}
