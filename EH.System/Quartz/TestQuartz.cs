using EH.Service.Interface.AD;
using EH.System.Commons;
using Quartz;

namespace EH.System.Quartz
{
    public class TestQuartz : IJob,ITransient
    {
        private readonly LogHelper logHelper;

        public TestQuartz(LogHelper logHelper)
        {
            this.logHelper = logHelper;
        }
        async Task IJob.Execute(IJobExecutionContext context)
        {
            logHelper.LogInfo("Test");
        }
    }
}
