using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Repository.Implement.Sys;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using Newtonsoft.Json;
using EH.Service.Interface.Sys;

namespace EH.Service.Implement.Sys
{
    public class SysMenuService : BaseService<Sys_Menus>, ISysMenuService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly ISysMenusRepository iSysMenusRepository;
        private readonly ISysRolesRepository roleRepository;
        private readonly ISysUserRoleRepository iSysUserRoleRepository;
        private readonly ISysUsersRepository iSysUsersRepository;
        private readonly ISysRoleMenuRepository iSysRoleMenuRepository;
        public SysMenuService(ISysUsersRepository iSysUsersRepository, ISysMenusRepository iSysMenusRepository, ISysRoleMenuRepository iSysRoleMenuRepository, ISysUserRoleRepository iSysUserRoleRepository, ISysRolesRepository roleRepository, LogHelper logHelper) : base(iSysMenusRepository, logHelper)
        {
            this.iSysUsersRepository = iSysUsersRepository;
            this.iSysMenusRepository = iSysMenusRepository;
            this.iSysRoleMenuRepository = iSysRoleMenuRepository;
            this.iSysUserRoleRepository = iSysUserRoleRepository;
            this.roleRepository = roleRepository;
        }


        public List<Menus> GetMenuListByUser(string userName)
        {
            var user = iSysUsersRepository.FirstOrDefault(x => x.UserName == userName);
            if (user is not null)
            {
                var userId = user.ID;
                var userRoles = iSysUserRoleRepository.Entities;
                var RoleMenus = iSysRoleMenuRepository.Entities;
                var menus = iSysMenusRepository.Entities;
                IQueryable<Sys_Menus> query;

                if (!iSysUsersRepository.isAdmin(userId))
                {
                    query = from ur in userRoles
                            join rm in RoleMenus on ur.RoleID equals rm.RoleID
                            join m in menus on rm.MenuID equals m.ID.ToString()
                            where ur.UserID == userId.ToString()
                            select m;
                }
                else
                {
                    //query = from ur in userRoles
                    //        join rm in RoleMenus on ur.RoleID equals rm.RoleID
                    //        join m in menus on rm.MenuID equals m.ID.ToString() 
                    //        select m;

                    query = from m in menus
                            select m;
                }
                var allMenuList = query.ToList();
                if (allMenuList.Any())
                {
                    var menuDto = new Menus();
                    var menuList = allMenuList.ToObject<List<Menus>>().OrderBy(i => i.Sort).ToList();
                    #region MyRegion
                    //var topMenuList = query.Where(m => string.IsNullOrEmpty(m.ParentID)).Select(e => new Menus()
                    //{
                    //    ID = e.ID,
                    //    Icon = e.Icon,
                    //    Status = e.Status,
                    //    Component = e.Component,
                    //    CreateBy = e.CreateBy,
                    //    CreateDate = e.CreateDate,
                    //    IsDeleted = e.IsDeleted,
                    //    MenuName = e.MenuName,
                    //    ModifyDate = e.ModifyDate,
                    //    MenuType = e.MenuType,
                    //    ModifyBy = e.ModifyBy,
                    //    ParentID = e.ParentID,
                    //    Permissions= e.Permissions,
                    //    Remark= e.Remark,
                    //    RouteName = e.RouteName,
                    //    RoutePath = e.RoutePath,
                    //    Visible = e.Visible,
                    //    Sort = e.Sort,
                    //});
                    #endregion
                    var newId = new Guid().ToString();
                    var topMenuList = allMenuList.Where(m => m.ParentID == newId).ToObject<List<Menus>>().OrderByDescending(i => i.Sort);
                    foreach (var topMenu in topMenuList)
                    {
                        topMenu.SetChildren(GetChildrenMenu(menuList, topMenu.ID.ToString()));
                    }
                    return topMenuList.ToList();
                }
                else
                {
                    return new List<Menus>();
                }
            }
            else
            {
                return new List<Menus>();

            }
        }

        public List<Sys_Menus> GetMenuByRole(string roleId)
        {

            var roles = roleRepository.Entities;
            var RoleMenus = iSysRoleMenuRepository.Entities;
            var menus = iSysMenusRepository.Entities;
            IQueryable<Sys_Menus> query; 
            query = from r in roles
                    join rm in RoleMenus on r.ID.ToString() equals rm.RoleID
                    join m in menus on rm.MenuID equals m.ID.ToString()
                    where r.ID.ToString() == roleId
                    orderby m.Sort ascending
                    select m;

            var allMenuList = query.ToList();
            if (allMenuList.Any())
            {
                var menuDto = new Sys_Menus();
                var menuList = allMenuList.ToObject<List<Sys_Menus>>().OrderBy(i => i.Sort).ToList();

                var newId = new Guid().ToString();
                var topMenuList = allMenuList.Where(m => m.ParentID == newId).ToObject<List<Sys_Menus>>();
                foreach (var topMenu in topMenuList)
                {
                    topMenu.SetChildren(GetChildrenMenu(menuList, topMenu.ID.ToString()));
                }
                return topMenuList.ToList();
            }
            else
            {
                return new List<Sys_Menus>();
            }


        }


        public List<Sys_Menus> GetParentMenuList(string userID)
        {

            //return GetMenuByRole(userID).Where(i => i.MenuType == 0).ToList();
            return GetAllMenus().Where(i => i.MenuType == 0).ToList();
        }

        public List<Sys_Menus> GetAllMenus()
        {
            var allMenus = iSysMenusRepository.Entities.ToList();
            var topMenuList = allMenus.Where(m => m.ParentID == new Guid().ToString()).ToList().ToObject<List<Sys_Menus>>();
            foreach (var topMenu in topMenuList)
            {
                topMenu.SetChildren(GetChildrenMenus(allMenus, topMenu.ID.ToString()));
            }
            return topMenuList.ToList();
        }




        private List<Sys_Menus> GetChildrenMenus(List<Sys_Menus> menuList, string menuId)
        {
            List<Sys_Menus> nextLevelMenuList = menuList.Where(menu => menu.ParentID == menuId).ToList();
            foreach (Sys_Menus nextLevelMenu in nextLevelMenuList)
            {
                nextLevelMenu.SetChildren(GetChildrenMenus(menuList, nextLevelMenu.ID.ToString()));
            }
            return nextLevelMenuList;
        }

        private List<Menus> GetChildrenMenu(List<Menus> menuList, string menuId)
        {
            List<Menus> nextLevelMenuList = menuList.Where(menu => menu.ParentID == menuId).ToList();
            foreach (Menus nextLevelMenu in nextLevelMenuList)
            {
                nextLevelMenu.SetChildren(GetChildrenMenu(menuList, nextLevelMenu.ID.ToString()));
            }
            return nextLevelMenuList;
        }

        private List<Sys_Menus> GetChildrenMenu(List<Sys_Menus> menuList, string menuId)
        {
            List<Sys_Menus> nextLevelMenuList = menuList.Where(menu => menu.ParentID == menuId).ToList();
            foreach (Sys_Menus nextLevelMenu in nextLevelMenuList)
            {
                nextLevelMenu.SetChildren(GetChildrenMenu(menuList, nextLevelMenu.ID.ToString()));
            }
            return nextLevelMenuList;
        }

        public List<Sys_Menus> GetMenuListByRole(string roleId)
        {
            var role = roleRepository.GetById(Guid.Parse(roleId));
            var roleMenus = iSysRoleMenuRepository.Entities;
            var menus = iSysMenusRepository.Entities;
            IQueryable<Sys_Menus> query;

            query = from rm in roleMenus
                    join m in menus on rm.MenuID equals m.ID.ToString()
                    where rm.RoleID == roleId
                    select m;


            var allMenuList = query.ToList();
            //if (allMenuList.Any())
            //{
            //    var topMenuList = allMenuList.Where(m => m.ParentID == new Guid().ToString()).ToList().ToObject<List<Sys_Menus>>();
            //    foreach (var topMenu in topMenuList)
            //    {
            //        topMenu.SetChildren(GetChildrenMenus(allMenuList, topMenu.ID.ToString()));
            //    }
            //    return topMenuList.ToList();
            //}
            //else
            //{
            //    return new List<Sys_Menus>();
            //}

            return allMenuList;
        }

    }
}
