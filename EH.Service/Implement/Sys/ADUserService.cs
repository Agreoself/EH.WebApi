using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Entities;
using EH.Repository.Interface.Sys;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EH.System.Models.Entities.Sys_Users;
using EH.Service.Interface.Sys;
using Microsoft.Extensions.Configuration;

namespace EH.Service.Implement.Sys
{
    public class ADUserService : IADUserService, ITransient
    {
        private readonly ISysUsersRepository sys_UsersRepository;
        private readonly IHttpContextAccessor httpContext;
        private readonly ADHelper aDHelper;
        private readonly IConfiguration configuration;
        private readonly LogHelper logHelper;

        public ADUserService(ISysUsersRepository sys_UsersRepository, IHttpContextAccessor httpContext, ADHelper aDHelper,IConfiguration configuration, LogHelper logHelper)
        {
            this.sys_UsersRepository = sys_UsersRepository;
            this.httpContext = httpContext;
            this.aDHelper = aDHelper;
            this.logHelper = logHelper;
            this.configuration = configuration;
        }

        public JsonResultModel<string> ResetPassword(string userName)
        {
            var res = aDHelper.ResetPassword(userName);
            return new JsonResultModel<string>()
            {
                Code = res != null ? "000" : "100",
                Message = res != null ? "重置成功" : "重置失败",
                Result = res,
                Other = null
            };
        }

        public JsonResultModel<bool> CheckLogin(string username, string password)
        {
            var res = aDHelper.CheckLogin(username, password);
            return new JsonResultModel<bool>()
            {
                Code = res ? "000" : "100",
                Message = res ? "登录成功" : "登录失败",
                Result = res,
                Other = null
            };
        }

        public JsonResultModel<object> GetUser(string userName)
        {
            var res = aDHelper.GetUserDNByName(userName);
            return new JsonResultModel<object>()
            {
                Code = res != null ? "000" : "100",
                Message = res != null ? "查询成功" : "查询失败",
                Result = res,
                Other = null
            };
        }

        public JsonResultModel<bool> GetGroupName()
        {
            //var res = aDHelper.GetXMUsers();
            return new JsonResultModel<bool>()
            {
               
            };
        }

        public JsonResultModel<bool> GenerateUser()
        {
            List<Sys_Users> sys_Users = new List<Sys_Users>();
            var results = aDHelper.GetAllUser();
            foreach (SearchResult result in results)
            {
                var userEntry = result.GetDirectoryEntry();

                Sys_Users user = new Sys_Users();
                user.ID = Guid.NewGuid();
                user.UserName = userEntry.Properties["samaccountname"].Value.ToString();
                user.FullName = userEntry.Properties["cn"].Value.ToString();
                user.Gender = 1;
                user.Department = userEntry.Properties["department"].Value == null ? "" : userEntry.Properties["department"].Value.ToString();
                user.Email = userEntry.Properties["mail"].Value == null ? "" : userEntry.Properties["mail"].Value.ToString();

                var jobtitle = userEntry.Properties["title"].Value == null ? "" : userEntry.Properties["title"].Value.ToString().ToLower();
                jobtitle = jobtitle.Contains("general manager") ? "vp" : jobtitle.Contains("director") ? "director" : jobtitle.Contains("supervisor") ? "supervisor" : jobtitle.Contains("manager") ? "manager" : "";
                user.JobTitle = jobtitle;

                var manager = userEntry.Properties["manager"].Value == null ? "" : userEntry.Properties["manager"].Value.ToString();
                if (manager != "")
                {
                    manager = manager.Substring(manager.IndexOf("="), manager.IndexOf(",") - manager.IndexOf("="));
                    manager = manager.Substring(1, manager.Length - 1);
                }
                user.Report = manager;

                var userAccountControl = userEntry.Properties["userAccountControl"].Value == null ? "" : userEntry.Properties["userAccountControl"].Value.ToString();
                int userActiveNum = Convert.ToInt32(userAccountControl);
                user.IsActive = (userActiveNum & 2) == 0;
                user.IsAdmin = false;

                user.EhiStratWorkDate = Convert.ToDateTime(userEntry.Properties["whencreated"].Value == null ? "" : userEntry.Properties["whencreated"].Value.ToString());

                user.Phone = userEntry.Properties["telephonenumber"].Value == null ? "" : userEntry.Properties["telephonenumber"].Value.ToString();
                user.CreateDate = DateTime.Now;
                user.CreateBy = httpContext.HttpContext.User.Identity.Name.ToString().Split("\\")[1];
                user.ModifyDate = Convert.ToDateTime(userEntry.Properties["whenchanged"].Value == null ? "" : userEntry.Properties["whenchanged"].Value.ToString());
                user.ModifyBy = httpContext.HttpContext.User.Identity.Name.ToString().Split("\\")[1];

                sys_Users.Add(user);
            }

            var nowUsers = sys_UsersRepository.TrackEntities.ToList();
            var updateUsers = nowUsers.Intersect(sys_Users, new UsersEqulityComparer()).ToList();
            var addUsers = sys_Users.Except(updateUsers);

            sys_UsersRepository.AddRange(addUsers, false);
            sys_UsersRepository.UpdateRange(updateUsers, isSave: false);
            var res = sys_UsersRepository.SaveChanges();

            return new JsonResultModel<bool>()
            {
                Code = res > 0 ? "000" : "100",
                Message = res > 0 ? "成功" : "失败",
                Result = res > 0,
                Other = null
            };

        }
    }
}
