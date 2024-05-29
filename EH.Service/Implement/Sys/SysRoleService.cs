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
using EH.Repository.Implement;
using NPOI.SS.Formula.Functions;
using EH.Service.Interface.Sys;
using System.Reflection; 

namespace EH.Service.Implement.Sys
{
    public class SysRoleService : BaseService<Sys_Roles>, ISysRoleService, ITransient
    {
        private readonly LogHelper logHelper;

        private readonly ISysRolesRepository iSysRolesRepository;
        private readonly ISysUserRoleRepository iSysUserRoleRepository;
        private readonly ISysUsersRepository iSysUsersRepository;
        private readonly ISysMenusRepository menuRepository;
        private readonly ISysRoleMenuRepository roleMenuRepository;

        public SysRoleService(ISysRolesRepository iSysRolesRepository, ISysUserRoleRepository iSysUserRoleRepository, ISysUsersRepository iSysUsersRepository, ISysMenusRepository menuRepository, ISysRoleMenuRepository roleMenuRepository, LogHelper logHelper) : base(iSysRolesRepository, logHelper)
        {
            this.iSysRolesRepository = iSysRolesRepository;
            this.iSysUserRoleRepository = iSysUserRoleRepository;
            this.iSysUsersRepository = iSysUsersRepository;
            this.menuRepository = menuRepository;
            this.roleMenuRepository = roleMenuRepository;
        }

        public List<Sys_Roles> GetAllRole()
        {
            return iSysRolesRepository.Entities.ToList();
        }

        public List<Sys_Roles> GetRoleByUser(string userId)
        {
            var userRoles = iSysUserRoleRepository.Entities;
            var roles = iSysRolesRepository.Entities;

            var query = from ur in userRoles
                        join r in roles on ur.RoleID equals r.ID.ToString()
                        where ur.UserID == userId
                        select r;

            var res = query.ToList();
            return res;
        }

        public List<Sys_Users> GetUserinfoInRole(string roleId)
        {
            var userRoles = iSysUserRoleRepository.Entities;
            var users = iSysUsersRepository.Entities;

            var query = from ur in userRoles
                        join u in users on ur.UserID equals u.ID.ToString()
                        where ur.RoleID == roleId
                        select u;

            var res = query.ToList();
            return res;
        }

        public bool DeleteUserRole(UserRole userRole)
        {
            try
            {
                var roleId = userRole.roleIds[0];
                var urList = iSysUserRoleRepository.Entities.Where(i => i.RoleID == roleId && userRole.userIds.Contains(i.UserID)).ToList();

                //List<Sys_UserRole> list = new List<Sys_UserRole>();
                //foreach (var id in ids)
                //{
                //    Sys_UserRole entity = iSysUserRoleRepository.GetById(Guid.Parse(id));
                //    if (entity != null)
                //    {
                //        //PropertyInfo propertyInfo = entity.GetType().GetProperty("IsDeleted");
                //        //if (propertyInfo != null && propertyInfo.PropertyType == typeof(bool))
                //        //{
                //        //    propertyInfo.SetValue(entity, true); // 修改 IsDeleted 属性为 true，实现软删除
                //        //    list.Add(entity);
                //        //}
                //        list.Add(entity);
                //    }
                //}

                iSysUserRoleRepository.DeleteRange(urList);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        public bool SetMenu(RoleMenu roleMenu)
        {
            try
            {
                var roles = iSysRolesRepository.Where(i => roleMenu.roleIds.Contains(i.ID.ToString())).ToList();
                var menus = menuRepository.Where(i => roleMenu.menuIds.Contains(i.ID.ToString())).ToList();

                List<Sys_RoleMenu> rms = new List<Sys_RoleMenu>();
                foreach (var role in roles)
                {
                    foreach (var menu in menus)
                    {
                        Sys_RoleMenu rm = new Sys_RoleMenu();
                        rm.RoleID = role.ID.ToString();
                        rm.MenuID = menu.ID.ToString();
                        rms.Add(rm);
                    }
                }

                var needDeleteRMs = roleMenuRepository.Entities.Where(rm => roleMenu.roleIds.Contains(rm.RoleID)).ToList();
                roleMenuRepository.DeleteRange(needDeleteRMs);
                roleMenuRepository.AddRange(rms);
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError(ex.ToString());
                return false;
            }
        

            //try
            //{
            //    var menus = menuRepository.Where(i => menuIds.Contains(i.ID.ToString())).ToList();
            //    var roles = iSysRolesRepository.Where(i => roleIds.Contains(i.ID.ToString())).ToList();
            //    var userRoles = iSysUserRoleRepository.Entities.ToList();

            //    List<Sys_RoleMenu> rms = new List<Sys_RoleMenu>();
            //    foreach (var role in roles)
            //    {
            //        foreach (var menu in menus)
            //        {
            //            Sys_RoleMenu rm = new Sys_RoleMenu();
            //            rm.RoleID = role.ID.ToString();
            //            rm.MenuID = menu.ID.ToString();
            //            rms.Add(rm);
            //        }
            //    }

            //    if (rms.Count > 0)
            //    {
            //        var existMenus = rms.Select(i => i.MenuID).ToList();
            //        var existRoles = rms.Select(i => i.RoleID).ToList();

            //        var existUrs = roleMenuRepository.Entities.Where(ur => existMenus.Contains(ur.MenuID) && existRoles.Contains(ur.RoleID)).ToList();

            //        roleMenuRepository.DeleteRange(existUrs);
            //        var addEntitys = rms.Except(existUrs);
            //        roleMenuRepository.AddRange(addEntitys);
            //    }
            //    else
            //    {
            //        var needDeleteUR = roleMenuRepository.Entities.Where(ur => menuIds.Contains(ur.MenuID)).ToList();
            //        roleMenuRepository.DeleteRange(needDeleteUR);
            //        roleMenuRepository.AddRange(rms);
            //    }
            //    return true;
            //}
            //catch (Exception ex)
            //{
            //    logHelper.LogError(ex.ToString());
            //    return false;
            //}
        }

    }
}
