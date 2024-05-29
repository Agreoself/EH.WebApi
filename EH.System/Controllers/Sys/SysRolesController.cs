using EH.System.Attribute;
using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Service.Implement;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using EH.Service.Interface.Sys;
using EH.Service.Interface;

namespace EH.System.Controllers.Sys
{
    [Route("api/[controller]")]
    [ApiController]
    public class SysRolesController : BaseController<Sys_Roles>
    {
        private readonly ISysRoleService roleService;
        public SysRolesController(ISysRoleService roleService) : base(roleService)
        {
            this.roleService = roleService;
        }

        [HttpGet]
        [Authorize]
        [Route("GetRole")]
        public JsonResultModel<List<Sys_Roles>> GetRole()
        { 
            var res = roleService.GetAllRole();
            return new JsonResultModel<List<Sys_Roles>>
            {
                Code = res != null ? "000" : "100",
                Result = res,
                Message = res != null ? "Success" : "False",
            };
        }

        [HttpPost]
        [Authorize]
        [Route("GetRoleByUser")]
        public JsonResultModel<List<Sys_Roles>> GetRoleByUser(string userId)
        {
            var res = roleService.GetRoleByUser(userId);
            return new JsonResultModel<List<Sys_Roles>>
            {
                Code = res != null ? "000" : "100",
                Result = res,
                Message = res != null ? "Success" : "False",
            };
        }

        [HttpPost]
        [Authorize]
        [Route("DeleteUserRole")]
        public virtual JsonResultModel<bool> DeleteUserRole(UserRole userRole)
        {
            var res = roleService.DeleteUserRole(userRole);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }

        [HttpPost]
        [Authorize]
        [Route("SetMenu")]
        public virtual JsonResultModel<bool> SetMenu(RoleMenu roleMenu)
        {
            var res = roleService.SetMenu(roleMenu);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }


    }
}
