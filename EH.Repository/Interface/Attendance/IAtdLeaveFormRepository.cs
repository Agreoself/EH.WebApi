using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Repository.Interface.Attendance
{
    public interface IAtdLeaveFormRepository : IRepositoryBase<Atd_LeaveForm>
    {
        //List<Atd_FormAndProcess> GetWaitAuditForm(PageRequest<Atd_LeaveForm> pageRequest, out int totalCount);
    }
}
