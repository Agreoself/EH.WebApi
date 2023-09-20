using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Repository.Implement.Sys;
using EH.Repository.Interface;
using EH.Repository.Interface.Sys;
using EH.Service.Interface;
using Newtonsoft.Json; 


namespace EH.Service.Implement
{
    public class SysMenuService : BaseService<Sys_Menus>, ISysMenuService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly ISysMenusRepository iSysMenusRepository;
        //private readonly ISysRolesRepository iSysRolesRepository;
        private readonly ISysUserRoleRepository iSysUserRoleRepository;
        private readonly ISysUsersRepository iSysUsersRepository;
        private readonly ISysRoleMenuRepository iSysRoleMenuRepository;
        public SysMenuService(ISysUsersRepository iSysUsersRepository, ISysMenusRepository iSysMenusRepository, ISysRoleMenuRepository iSysRoleMenuRepository, ISysUserRoleRepository iSysUserRoleRepository,LogHelper logHelper) : base(iSysMenusRepository, logHelper)
        {
            this.iSysUsersRepository = iSysUsersRepository;
            this.iSysMenusRepository = iSysMenusRepository;
            this.iSysRoleMenuRepository = iSysRoleMenuRepository;
            this.iSysUserRoleRepository = iSysUserRoleRepository;
        }
         
    
        public List<Menus> GetMenuListByUser(string userName)
        {
            var user = iSysUsersRepository.FirstOrDefault(x => x.UserName == userName);
            if (user is not null)
            {
                var userId=user.ID;
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

                    query=from m in menus 
                          select m;
                }
                var allMenuList=query.ToList();
                if (allMenuList.Any())
                {
                    var menuDto = new Menus();
                    var menuList = allMenuList.ToObject<List<Menus>>();
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

                    var topMenuList = allMenuList.Where(m => string.IsNullOrEmpty(m.ParentID)).ToObject<List<Menus>>();
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

       
        public List<Menus> GetParentMenuList(string userID)
        {
            return GetMenuListByUser(userID).Where(i => i.MenuType == 0).ToList();
        }
         
        public List<Sys_Menus> GetAllMenus()
        {
            var allMenus = iSysMenusRepository.Entities.ToList();
            var topMenuList = allMenus.Where(m => string.IsNullOrEmpty(m.ParentID)).ToList().ToObject<List<Sys_Menus>>();
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
    }
}
