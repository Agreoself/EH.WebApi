using EH.System.Attribute;
using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Service.Implement;
using EH.Service.Interface;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EH.System.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SysRolesController : BaseController<Sys_Roles>
    {
        private readonly ISysRoleService roleService;
        public SysRolesController(ISysRoleService roleService):base(roleService)
        {
            this.roleService = roleService;
        }
         
        [HttpGet]
        [Authorize]
        [Route("GetRole")]
        public JsonResultModel<Sys_Roles> GetRole()
        {
            var userName = HttpContext.User.Identity!.Name!.Split('\\')[1];
            var res = new Sys_Roles();// roleService(userName);
            return new JsonResultModel<Sys_Roles>
            {
                Code = res!=null ? "000" : "100",
                Result = res,
                Message = res != null ? "Success" : "False",
            };
        }

  

    }
}
