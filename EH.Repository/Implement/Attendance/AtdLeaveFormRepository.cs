using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.DataAccess;
using EH.Repository.Interface.Attendance;
using EH.System.Models.Dtos;
using EH.Repository.Interface.Sys;

namespace EH.Repository.Implement.Sys
{
    public class AtdLeaveFormRepository : RepositoryBase<Atd_LeaveForm>, IAtdLeaveFormRepository, ITransient
    {
        //private readonly IAtdLeaveProcessRepository processRepository;
        //private readonly ISysUsersRepository usersRepository;

        public AtdLeaveFormRepository(MyAppDbContext mydbcontext) : base(mydbcontext)
        {

        }

        //public List<Atd_FormAndProcess> GetWaitAuditForm(PageRequest<Atd_LeaveForm> pageRequest, out int totalCount)
        //{
        //    var userId = pageRequest.defaultWhere.Split(',')[0].Split('=')[1];
        //    var user = usersRepository.Entities.FirstOrDefault(i => i.UserName == userId);
        //    var processEntity = processRepository.Entities.Where(i => user.FullName.Contains(i.UserId) || user.UserName == i.UserId).ToList();
        //    List<Atd_FormAndProcess> processForms = new List<Atd_FormAndProcess>();
        //    foreach (var process in processEntity)
        //    {
        //        var leaveId = process.LeaveId;
        //        var form = base.FirstOrDefault(i => i.LeaveId == leaveId && process.ProcessState == "wait");
        //        var processForm = form.ToObject<Atd_FormAndProcess>();
        //        if (form != null)
        //        {
        //            processForm.ProcessID = process.ID.ToString();
        //            processForm.FormID = form.ID.ToString();
        //            processForms.Add(processForm);
        //        }
        //    }
        //    totalCount = processForms.Count;
        //    var res = processForms.Skip((pageRequest.PageIndex - 1) * pageRequest.PageSize).Take(pageRequest.PageSize);
        //    return processForms;
        //}
    }
}
