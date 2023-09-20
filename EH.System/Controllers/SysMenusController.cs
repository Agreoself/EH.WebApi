using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Service.Interface;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace EH.System.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SysMenusController : BaseController<Sys_Menus>
    {
        private readonly ISysMenuService sysMenuService;
        public SysMenusController(ISysMenuService sysMenuService):base(sysMenuService) 
        {
            this.sysMenuService = sysMenuService;
        }

        [HttpGet]
        [Authorize]
        [Route("GetMenuListByUser")]
        public JsonResultModel<List<Menus>> GetMenuListByUser()
        {
            var userName = HttpContext.User.Identity!.Name!.Split('\\')[1];
            var menuList = sysMenuService.GetMenuListByUser(userName);
            return new JsonResultModel<List<Menus>> {
                Code = menuList.Any() ? "000" : "100",
                Message = menuList.Any()?"success":"fail",
                Result = menuList
            }; 
        }

        [HttpGet]
        [Authorize]
        [Route("GetParentMenuList")]
        public JsonResultModel<List<Menus>> GetParentMenuList()
        {
            var userName = HttpContext.User.Identity!.Name!.Split('\\')[1];
            var menuList = sysMenuService.GetParentMenuList(userName);
            return new JsonResultModel<List<Menus>>
            {
                Code = menuList.Any() ? "000" : "100",
                Message = menuList.Any() ? "success" : "fail",
                Result = menuList
            };
        }

        [HttpGet]
        [Authorize]
        [Route("GetAllMenus")]
        public JsonResultModel<List<Sys_Menus>> GetAllMenus()
        {
            var menuList = sysMenuService.GetAllMenus();
            return new JsonResultModel<List<Sys_Menus>>
            {
                Code = menuList.Any() ? "000" : "100",
                Message = menuList.Any() ? "success" : "fail",
                Result = menuList
            };
        }


    }
}

