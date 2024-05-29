using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Repository.Implement.Sys;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using EH.Service.Interface.Sys;
using EH.Repository.Interface.Attendance;
using EH.Service.Interface.Attendance;
using NPOI.POIFS.FileSystem;

namespace EH.Service.Implement.Sys
{
    public class SysUserService : BaseService<Sys_Users>, ISysUserService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IADUserService iADUserService;
        private readonly ISysUsersRepository iSysUsersRepository;
        private readonly ISysRolesRepository iSysRolesRepository;
        private readonly ISysUserRoleRepository iSysUserRoleRepository;
        private readonly IAtdLeaveFormRepository formRepository;
        private readonly IAtdLeaveFormService formService;
        private readonly IAtdLeaveProcessRepository processRepository;
        public SysUserService(IADUserService iADUserService,
            ISysUsersRepository iSysUsersRepository, ISysRolesRepository iSysRolesRepository, ISysUserRoleRepository iSysUserRoleRepository, IAtdLeaveFormRepository formRepository, IAtdLeaveProcessRepository processRepository, IAtdLeaveFormService formService, LogHelper logHelper) : base(iSysUsersRepository, logHelper)
        {
            this.iADUserService = iADUserService;
            this.iSysUsersRepository = iSysUsersRepository;
            this.iSysRolesRepository = iSysRolesRepository;
            this.iSysUserRoleRepository = iSysUserRoleRepository;
            this.formRepository = formRepository;
            this.processRepository = processRepository;
            this.logHelper = logHelper;
            this.formService = formService;
        }

        public UserInfo GetUserInfo(string userName)
        {
            UserInfo userInfo = new UserInfo();

            var user = iSysUsersRepository.FirstOrDefault(i => i.UserName == userName);
            if (user is not null)
            {
                Guid userID = user.ID;
                userInfo = JsonConvert.DeserializeObject<UserInfo>(JsonConvert.SerializeObject(user));
                Dictionary<string, string> roleList = new Dictionary<string, string>();
                var users = iSysUsersRepository.Entities;
                var roles = iSysRolesRepository.Entities;
                var userRoles = iSysUserRoleRepository.Entities;
                var query = from u in users
                            join ur in userRoles on u.ID.ToString() equals ur.UserID
                            join r in roles on ur.RoleID equals r.ID.ToString()
                            where u.ID == userID
                            select r;
                foreach (var u in query)
                {
                    roleList.Add(u.RoleName, u.ID.ToString());
                }
                userInfo.RoleList = roleList;
                #region MyRegion
                //select new UserInfo
                //{
                //    UserName = u.UserName,
                //    FullName = u.FullName,
                //    Gender = u.Gender,
                //    Department = u.Department,
                //    Email = u.Email,
                //    IsAdmin = u.IsAdmin,
                //    IsActive = u.IsActive,
                //    JobTitle = u.JobTitle,
                //    StartWorkDate = u.StartWorkDate,
                //    EHIStratWorkDate = u.EHIStratWorkDate,
                //    CC = u.CC,
                //    Phone = u.Phone,
                //    Report = u.Report,
                //    RoleList = new Dictionary<string, string>
                //    {
                //        { ur.RoleID,r.RoleName }
                //    }
                //};
                #endregion

                return userInfo;
            }
            else
            {
                return null;
            }

        }

        public bool Login(User loginForm)
        {
            var user = iSysUsersRepository.FirstOrDefault(i => i.UserName.ToLower() == loginForm.username.ToLower());
            if (user == null)
            {
                return false;
            }
            if (loginForm.noPassword)
            {
                var users = iSysUsersRepository.Entities;
                var roles = iSysRolesRepository.Entities;
                var userRoles = iSysUserRoleRepository.Entities;
                var query = from u in users
                            join ur in userRoles on u.ID.ToString() equals ur.UserID
                            join r in roles on ur.RoleID equals r.ID.ToString()
                            where u.UserName == loginForm.orginUser
                            select r;
                var isAdminOrHr = false;
                var HrAndSu = new List<string> { "超级管理员", "HR" };
                foreach (var q in query)
                {
                    if (HrAndSu.Contains(q.RoleName))
                    {
                        isAdminOrHr=true;break;
                    }
                }
                return isAdminOrHr; //判断是否管理员或者hr，如果不是则没有权限免密登录
            }
            else
            {
                if (!iADUserService.CheckLogin(loginForm.username, loginForm.password).Result)
                {
                    return false;
                }
                return true;
            }

            //if (string.IsNullOrEmpty(user.Password))
            //{
            //    user.Password = (loginForm.password);
            //    iSysUsersRepository.Update(user);
            //}
            //string token = GenerateToken(user);

        }

        public List<Sys_Users> GetUserListInRole(PageRequest<Sys_Users> request, out int totalCout)
        {
            var userRoles = iSysUserRoleRepository.Entities;
            var users = iSysUsersRepository.Entities;
            var roleId = request.defaultWhere.Split('=')[1];

            var where = request.GetWhere().Compile();

            var query = from ur in userRoles
                        join u in users on ur.UserID equals u.ID.ToString()
                        where ur.RoleID == roleId
                        orderby u.CreateDate descending
                        select u;

            var result = query.Where(where);
            totalCout = result.Count();

            var res = result.Skip((request.PageIndex - 1) * request.PageSize).Take(request.PageSize).ToList();
            return res;
        }

        public bool GrantRole(List<string> userIds, List<string> roleIds)
        {
            try
            {
                var users = iSysUsersRepository.Where(i => userIds.Contains(i.ID.ToString())).ToList();
                var roles = iSysRolesRepository.Where(i => roleIds.Contains(i.ID.ToString())).ToList();
                var userRoles = iSysUserRoleRepository.Entities.ToList();

                List<Sys_UserRole> urs = new List<Sys_UserRole>();
                foreach (var role in roles)
                {
                    foreach (var user in users)
                    {
                        Sys_UserRole ur = new Sys_UserRole();
                        ur.RoleID = role.ID.ToString();
                        ur.UserID = user.ID.ToString();
                        urs.Add(ur);
                    }
                }

                if (urs.Count > 0)
                {
                    var existUsers = urs.Select(i => i.UserID).ToList();
                    var existRoles = urs.Select(i => i.RoleID).ToList();

                    var existUrs = iSysUserRoleRepository.Entities.Where(ur => existUsers.Contains(ur.UserID) && existRoles.Contains(ur.RoleID)).ToList();

                    iSysUserRoleRepository.DeleteRange(existUrs);
                    var addEntitys = urs.Except(existUrs);
                    iSysUserRoleRepository.AddRange(addEntitys);
                }
                else
                {
                    var needDeleteUR = iSysUserRoleRepository.Entities.Where(ur => userIds.Contains(ur.UserID)).ToList();
                    iSysUserRoleRepository.DeleteRange(needDeleteUR);
                    iSysUserRoleRepository.AddRange(urs);
                }
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError(ex.ToString());
                return false;
            }
        }


        public bool SetRole(List<string> userIds, List<string> roleIds)
        {
            try
            {
                var users = iSysUsersRepository.Where(i => userIds.Contains(i.ID.ToString())).ToList();
                var roles = iSysRolesRepository.Where(i => roleIds.Contains(i.ID.ToString())).ToList();
                var userRoles = iSysUserRoleRepository.Entities.ToList();

                List<Sys_UserRole> urs = new List<Sys_UserRole>();
                foreach (var role in roles)
                {
                    foreach (var user in users)
                    {
                        Sys_UserRole ur = new Sys_UserRole();
                        ur.RoleID = role.ID.ToString();
                        ur.UserID = user.ID.ToString();
                        urs.Add(ur);
                    }
                }

                var needDeleteUR = iSysUserRoleRepository.Entities.Where(ur => userIds.Contains(ur.UserID)).ToList();
                iSysUserRoleRepository.DeleteRange(needDeleteUR);
                iSysUserRoleRepository.AddRange(urs);

                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError(ex.ToString());
                return false;
            }
        }

        public override bool Update(Sys_Users entity)
        {
            var user = iSysUsersRepository.FirstOrDefault(i => i.ID == entity.ID);
            var oldReport = user.Report;
            var newReport = entity.Report;
            if (oldReport != newReport)
            {
                var res = base.Update(entity);
                if (res)
                {
                    var reportUser = iSysUsersRepository.FirstOrDefault(i => i.UserName == newReport || i.FullName.Contains(newReport));
                    var forms = formRepository.Where(i => i.UserId == entity.UserName && i.CurrentState > -1 && i.CurrentState < 2).ToList();
                    foreach (var form in forms)
                    {
                        var process = processRepository.FirstOrDefault(i => i.LeaveId == form.LeaveId && i.UserId == oldReport && i.Action == "audit");
                        if (process != null)
                        {
                            process.UserId = newReport;
                            processRepository.Update(process);
                            //formService.SendMailByCurrentState(entity, reportUser, null, form);
                        }
                    }
                }
                return res;
            }
            else
            {
                var res = base.Update(entity);
                return res;
            }

        }

        public bool ChangePhoto(UserAvatar userAvatar)
        {
            var user = iSysUsersRepository.FirstOrDefault(i => i.ID == userAvatar.ID||i.UserName==userAvatar.userId);
            user.Avatar = userAvatar.Avatar;
            var res = base.Update(user);
            return res;
        }
    }
}
