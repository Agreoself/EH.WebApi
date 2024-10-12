using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EH.System.Commons.LeaveBalanceHelper;

namespace EH.Service.Interface.Attendance
{
    public interface IAtdLeaveBalanceService : IBaseService<Atd_LeaveBalance>
    {
        //bool CalculateAnnualAndSick(string userId);
        bool CalculateAnnualAndSick(List<string> userIds);

        bool CalculatePersonal(List<string> userIds);
        bool CalculateParentalAndBreastfeeding(string userId, DateTime bornDate);

        List<Atd_AnnualInfos> GetInfo(Atd_AnnualInfoReq data);

        List<Atd_BalanceStatics> Statistics(PageRequest request,out int totalCount);

        public string ClearCarryoverAnnual(bool isClear=false);
    }
}
