using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using NPOI.SS.Formula.Eval;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EH.Service.Implement.Attendance.AtdLeaveFormService;

namespace EH.Service.Interface.Attendance
{
    public interface IAtdLeaveFormService : IBaseService<Atd_LeaveForm>
    {
        List<Atd_FormWithState> GetPageListWithState(PageRequest<Atd_LeaveForm> request, out int totalCount);
        List<Atd_FormWithState> QueryGetPageList(PageRequest<Atd_LeaveForm> request, out int totalCount);
        List<Atd_FormAndProcess> GetWaitAuditForm(PageRequest<Atd_LeaveForm> pageRequest, out int totalCount);
        Atd_LeaveForm Save(Atd_LeaveForm T);
        JsonResultModel<Atd_LeaveForm> Apply(Atd_LeaveForm t);

        JsonResultModel<Atd_LeaveForm> UpdateFP(Atd_LeaveForm entity);
        JsonResultModel<Atd_LeaveForm> UploadAttachment(Atd_LeaveForm entity);
        bool AuditForm(Atd_Audit t);

        //JsonResultModel<Atd_LeaveForm> AuditForm(Atd_Audit t);

        bool SendMail(string formEmail, string toEmail, string subject, string body);

        List<Atd_Statics> GetStatistic(PageRequest request, out int totalCout);

        Dictionary<string, int> GetHomePageData(string userId);
        List<decimal> GetHomePageBodyData(string userId);
        void SendMailByCurrentState(Sys_Users user, Sys_Users auditUser, Sys_Users nextUser, Atd_LeaveForm fEntity);

        bool Treate(List<string> ids);

        double CalculateLeaveHours(DateTime dtStart, DateTime dtEnd, string workTime, bool isContainHoliday,bool isNursing);
        bool ApproveByEmail(Atd_ApproveByEmail approveByEmail);

        public void updateAttachment();
    }
}
