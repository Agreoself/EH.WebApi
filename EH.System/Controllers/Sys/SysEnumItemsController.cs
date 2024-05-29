using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using EH.Service.Interface.Sys;

namespace EH.System.Controllers.Sys
{
    [Route("api/[controller]")]
    [ApiController]
    public class SysEnumItemsController : BaseController<Sys_EnumItem>
    {
        private readonly ISysEnumItemService service;
        public SysEnumItemsController(ISysEnumItemService service) : base(service)
        {
            this.service = service;
        }

    }
}

