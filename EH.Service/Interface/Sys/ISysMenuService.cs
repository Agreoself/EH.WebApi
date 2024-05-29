using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.Sys
{
    public interface ISysMenuService : IBaseService<Sys_Menus>
    {

        public List<Menus> GetMenuListByUser(string userID); 
        public List<Sys_Menus> GetParentMenuList(string userID);
        public List<Sys_Menus> GetAllMenus();
        public List<Sys_Menus> GetMenuListByRole(string roleId);
        public List<Sys_Menus> GetMenuByRole(string roleId);
    }
}
