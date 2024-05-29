using EH.System.Commons;
using EH.System.Models.Common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using EH.Service.Interface.Sys;
using EH.Service.Interface.AD;
using EH.System.Models.Entities;
using EH.System.Models.Dtos;

namespace EH.System.Controllers.AD
{
    [Route("api/[controller]")]
    [ApiController]
    public class ADUserOperateController : ControllerBase
    {
        private readonly ILogger<ADUserOperateController> logger;
        private readonly IADUserOperateService adService;
        public ADUserOperateController(ILogger<ADUserOperateController> logger, IADUserOperateService adService)
        {
            this.logger = logger;
            this.adService = adService;
        }
  

        [HttpGet]
        [Authorize]
        [Route("GetUserByUserName")]
        public JsonResultModel<Sys_ADUsers> GetUserByUserName(string userName)
        {
            var res = adService.GetUserByUserName(userName);
            return new JsonResultModel<Sys_ADUsers>
            {
                Code = res!=null ? "000" : "100",
                Result = res ,
                Message = res != null ? "Success" : "Fail",
            };
        }
  


    }
}
