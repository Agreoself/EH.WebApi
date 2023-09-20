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
using EH.Repository.Implement;
using NPOI.SS.Formula.Functions;

namespace EH.Service.Implement
{
    public class SysRoleService : BaseService<Sys_Roles>, ISysRoleService, ITransient
    {
        private readonly LogHelper logHelper;

        private readonly ISysRolesRepository iSysRolesRepository;
        private readonly ISysUserRoleRepository iSysUserRoleRepository;
        private readonly ISysUsersRepository iSysUsersRepository;
        public SysRoleService(ISysRolesRepository iSysRolesRepository, ISysUserRoleRepository iSysUserRoleRepository, ISysUsersRepository iSysUsersRepository, LogHelper logHelper) : base(iSysRolesRepository, logHelper)
        {
            this.iSysRolesRepository = iSysRolesRepository;
            this.iSysUserRoleRepository = iSysUserRoleRepository;
            this.iSysUsersRepository = iSysUsersRepository;
        }

        public bool GrantRole(List<string> userIds)
        {
            return true;
        }

        public List<Sys_Users> GetUserinfoInRole(string roleId)
        {
            var userRoles = iSysUserRoleRepository.Entities;
            var users = iSysUsersRepository.Entities;
            
            var query= from ur in userRoles
                       join u in users on ur.UserID equals u.ID.ToString()
                       where ur.RoleID == roleId
                       select u; 
               
            var res= query.ToList();
            return res;
        }

    }
}
