using EH.Service.Interface.AD;
using EH.System.Commons;
using Quartz;

namespace EH.System.Quartz
{
    public class PasswordExpiredNotify : IJob, ITransient
    {
        private readonly IADUserPwdNotifyService userPwdNotifyService;
        private readonly LogHelper logHelper;
        private readonly EmailService emailService;
        private readonly IConfiguration configuration;

        public PasswordExpiredNotify(IADUserPwdNotifyService _userPwdNotifyService, LogHelper logHelper, EmailService emailService, IConfiguration configuration)
        {
            this.userPwdNotifyService = _userPwdNotifyService;
            this.logHelper = logHelper;
            this.emailService = emailService;
            this.configuration = configuration;
        }
        Task IJob.Execute(IJobExecutionContext context)
        {
            logHelper.LogInfo("start notify password expired user" + DateTime.Now.ToString());
            var notifyLocation = configuration.GetSection("ADSetting:ExpireLocationNotify").GetChildren().Select(i => i.Value).ToArray();
            foreach (var location in notifyLocation)
            {
                var needNotifyUsers = userPwdNotifyService.GenerateADUser(location);
                logHelper.LogInfo(location + " need Notify User count:" + needNotifyUsers.Count);
                logHelper.LogInfo(location + " need Notify User Info :" + needNotifyUsers.ToJson());
                userPwdNotifyService.SendMail(needNotifyUsers);
            }
            logHelper.LogInfo("compelted notify password expired user" + DateTime.Now.ToString());
            return Task.CompletedTask;
        }
    }
}
