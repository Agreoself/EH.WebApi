﻿using EH.System.Attribute;
using EH.System.Commons;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using EH.Service.Implement;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using EH.Service.Interface.Sys;

namespace EH.System.Controllers.Sys
{
    [Route("api/[controller]")]
    [ApiController]
    public class SysUsersController : BaseController<Sys_Users>
    {
        private readonly ISysUserService userService;
        public SysUsersController(ISysUserService sysUserService) : base(sysUserService)
        {
            userService = sysUserService;
        }

        [HttpPost]
        //[AllowAnonymous]
        [Route("Login")]
        public JsonResultModel<bool> Login(User user)
        {
            var orginUser = HttpContext.User.Identity!.Name!.Split('\\')[1];
            user.orginUser = orginUser;
            var res = userService.Login(user);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Result = res,
                Message = res ? "Success" : "False",
            };
        }

        [HttpGet]
        [Authorize]
        [Route("GetUserInfo")]
        public JsonResultModel<UserInfo> GetUserInfo(string userName)
        {
            //var userName = HttpContext.User.Identity!.Name!.Split('\\')[1];
            var res = userService.GetUserInfo(userName);
            return new JsonResultModel<UserInfo>
            {
                Code = res != null ? "000" : "100",
                Result = res,
                Message = res != null ? "Success" : "False",
            };
        }

        //[HttpPost]
        //[Authorize]
        //[Route("GetUserPageList")]
        //public JsonResultModel<List<Sys_Users>> GetUserPageList([FromBody] PageRequest<Sys_Users> pageRequest)
        //{
        //    var userName = HttpContext.User.Identity.Name.Split('\\')[1];

        //    var userList = _sysUserService.GetPageList(pageRequest, out int total);
        //    return new JsonResultModel<List<Sys_Users>>
        //    {
        //        Code = "000",
        //        Message = "success",
        //        Result = userList,
        //        Other = total,
        //    };
        //}

        [HttpPost]
        [Authorize]
        [Route("GrantRole")]
        public JsonResultModel<bool> GrantRole(UserRole userRole)
        {
            var res = userService.GrantRole(userRole.userIds, userRole.roleIds);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }

        [HttpPost]
        [Authorize]
        [Route("SetRole")]
        public JsonResultModel<bool> SetRole(UserRole userRole)
        {
            var res = userService.SetRole(userRole.userIds, userRole.roleIds);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }

        [HttpPost]
        [Authorize]
        [Route("GetUserListInRole")]
        public JsonResultModel<List<Sys_Users>> GetUserinfoInRole(PageRequest<Sys_Users> request)
        {
            var res = userService.GetUserListInRole(request, out int total);
            return new JsonResultModel<List<Sys_Users>>
            {
                Code = res != null ? "000" : "100",
                Result = res,
                Message = res != null ? "Success" : "False",
                Other = total
            };
        }

        [HttpPost]
        [Authorize]
        [Route("ChangePhoto")]
        public JsonResultModel<bool> ChangePhoto(UserAvatar dto)
        {
            var res = userService.ChangePhoto(dto);
            return new JsonResultModel<bool>
            {
                Code = res != null ? "000" : "100",
                Result = res,
                Message = res != null ? "Success" : "False",
            };
        }
    }
}
