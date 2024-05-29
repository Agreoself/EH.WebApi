using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using EH.Service.Implement;
using NPOI.POIFS.FileSystem;
using EH.Service.Interface.Sys;

namespace EH.System.Controllers.Sys
{
    [Route("api/[controller]")]
    [ApiController]
    public class SysEnumTypesController : BaseController<Sys_EnumType>
    {
        private readonly ISysEnumTypeService service;
        public SysEnumTypesController(ISysEnumTypeService service) : base(service)
        {
            this.service = service;
        }

        [HttpGet]
        [Route("GetAllDic")]
        [Authorize]
        public JsonResultModel<Dictionary<string, object>> GetAllDic()
        {
            var res = service.GetAllDic();
            return new JsonResultModel<Dictionary<string, object>>
            {
                Code = res is not null ? "000" : "100",
                Message = res is not null ? "success" : "fail",
                Result = res
            };
        }
    }
}

