using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Repository.Implement.Sys;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using EH.Service.Interface;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace EH.Service.Implement
{
    public class SysUserService : BaseService<Sys_Users>, ISysUserService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IADUserService iADUserService;
        private readonly ISysUsersRepository iSysUsersRepository;
        private readonly ISysRolesRepository iSysRolesRepository;
        private readonly ISysUserRoleRepository iSysUserRoleRepository;
        public SysUserService(IADUserService iADUserService,
            ISysUsersRepository iSysUsersRepository, ISysRolesRepository iSysRolesRepository, ISysUserRoleRepository iSysUserRoleRepository, LogHelper logHelper) : base(iSysUsersRepository, logHelper)
        {
            this.iADUserService = iADUserService;
            this.iSysUsersRepository = iSysUsersRepository;
            this.iSysRolesRepository = iSysRolesRepository;
            this.iSysUserRoleRepository = iSysUserRoleRepository;
        }

        public UserInfo GetUserInfo(string userName)
        {
            UserInfo userInfo = new UserInfo();

            var user = iSysUsersRepository.FirstOrDefault(i => i.UserName == userName);
            if (user is not null)
            {
                Guid userID = user.ID;
                userInfo = JsonConvert.DeserializeObject<UserInfo>(JsonConvert.SerializeObject(user));
                List<string> roleList = new List<string>();
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
                    roleList.Add(u.ID.ToString());
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
            if (!iADUserService.CheckLogin(loginForm.username, loginForm.password).Result)
            {
                return false;
            }
            if (string.IsNullOrEmpty(user.Password))
            {
                user.Password = HashHelper.Md5(loginForm.password);
                iSysUsersRepository.Update(user);
            }
            //string token = GenerateToken(user);
            return true;
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

                var existUsers = urs.Select(i => i.UserID).ToList();
                var existRoles = urs.Select(i => i.RoleID).ToList();

                var existUrs = iSysUserRoleRepository.Entities.Where(ur => existUsers.Contains(ur.UserID) && existRoles.Contains(ur.RoleID)).ToList();

                iSysUserRoleRepository.DeleteRange(existUrs);

                iSysUserRoleRepository.AddRange(urs.Except(existUrs));
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError(ex.ToString());
                return false;
            }
        }



    }
}
