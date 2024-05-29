using EH.System.Commons;
using EH.System.Models.Entities;
using Microsoft.AspNetCore.Http;
using System.DirectoryServices;
using EH.Service.Interface.AD;
using Microsoft.Extensions.Configuration;
using EH.Repository.Interface.AD;
using System.DirectoryServices.AccountManagement;
using NLog;
using ActiveDs;
using NPOI.XWPF.UserModel;
using EH.System.Models.Dtos;

namespace EH.Service.Implement.AD
{
    public class ADUserOperateService : BaseService<Sys_Users>, IADUserOperateService, ITransient
    { 
        private readonly IHttpContextAccessor httpContext;
        private readonly ADHelper aDHelper;
        private readonly IConfiguration configuration;
        private readonly LogHelper logHelper;
        private readonly EmailService emailService;

        public ADUserOperateService(IHttpContextAccessor httpContext, ADHelper aDHelper, IConfiguration configuration, LogHelper logHelper, EmailService emailService)
        { 
            this.httpContext = httpContext;
            this.aDHelper = aDHelper;
            this.logHelper = logHelper;
            this.configuration = configuration;
            this.emailService = emailService;
        }

        public Sys_ADUsers GetUserByUserName(string userName)
        {
            Sys_ADUsers user = new Sys_ADUsers();
            var userEntry = aDHelper.GetUserDNByUserName(userName);
            if (userEntry != null)
            {
                user.UserName = userEntry.Properties["samaccountname"][0].ToString();
                user.FullName = userEntry.Properties["cn"][0].ToString();
                user.Gender = 1;
                user.Department = userEntry.Properties["department"][0] == null ? "" : userEntry.Properties["department"][0].ToString();
                user.Email = userEntry.Properties["mail"][0] == null ? "" : userEntry.Properties["mail"][0].ToString();

                var jobtitle = userEntry.Properties["title"][0] == null ? "" : userEntry.Properties["title"][0].ToString().ToLower();
                jobtitle = jobtitle.Contains("general manager") ? "vp" : jobtitle.Contains("director") ? "director" : jobtitle.Contains("supervisor") ? "supervisor" : jobtitle.Contains("manager") ? "manager" : "";
                user.JobTitle = jobtitle;

                var manager = userEntry.Properties["manager"].Count<=0? "" : userEntry.Properties["manager"][0].ToString();
                if (manager != "")
                {
                    manager = manager.Substring(manager.IndexOf("="), manager.IndexOf(",") - manager.IndexOf("="));
                    manager = manager.Substring(1, manager.Length - 1);
                }
                user.Report = manager;

                var userAccountControl = userEntry.Properties["userAccountControl"].Count<=0 ? "0" : userEntry.Properties["userAccountControl"][0].ToString();
                int userActiveNum = Convert.ToInt32(userAccountControl);
                user.IsActive = (userActiveNum & 2) == 0;
                user.IsAdmin = false;
                user.EhiStratWorkDate = Convert.ToDateTime(userEntry.Properties["whencreated"].Count<=0? DateTime.Now : userEntry.Properties["whencreated"][0].ToString());
                return user;
            }
            else
            {
                return null;
            }
        }

      
    }
}
