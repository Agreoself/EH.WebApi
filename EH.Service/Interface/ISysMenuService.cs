using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface
{
    public interface ISysMenuService : IBaseService<Sys_Menus>
    {

        public List<Menus> GetMenuListByUser(string userID); 
        public List<Menus> GetParentMenuList(string userID);
        public List<Sys_Menus> GetAllMenus();

    }
}
