using EH.Service.Interface.Attendance;
using EH.System.Commons;
using Quartz;

namespace EH.System.Quartz
{
    public class UnapprovedNotify : IJob, ITransient
    {
        private readonly IAtdLeaveProcessService processService;
        private readonly LogHelper logHelper;

        public UnapprovedNotify(IAtdLeaveProcessService processService, LogHelper logHelper)
        {
            this.processService = processService;
            this.logHelper = logHelper;
        }
        public Task Execute(IJobExecutionContext context)
        {
            logHelper.LogInfo("start notify UnapprovedUser" + DateTime.Now.ToString());
            processService.GetUnapprovedUser();
            logHelper.LogInfo("compelted notify UnapprovedUser" + DateTime.Now.ToString());
            return Task.CompletedTask;
        }
    }
}
