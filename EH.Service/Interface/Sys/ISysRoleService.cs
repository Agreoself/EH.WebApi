using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.Sys
{
    public interface ISysRoleService : IBaseService<Sys_Roles>
    {
        List<Sys_Roles> GetAllRole();

        List<Sys_Roles> GetRoleByUser(string userId);
        bool DeleteUserRole(UserRole userRole);

        bool SetMenu(RoleMenu roleMenu);
    }
}
