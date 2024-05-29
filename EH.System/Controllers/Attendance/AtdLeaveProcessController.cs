using EH.System.Commons;
using EH.System.Models.Common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using EH.System.Models.Entities;
using EH.Service.Interface.Attendance;

namespace EH.System.Controllers.Attendance
{
    [Route("api/[controller]")]
    [ApiController]
    public class AtdLeaveProcessController : BaseController<Atd_LeaveProcess>
    {
        private readonly ILogger<AtdLeaveProcessController> logger;
        private readonly IAtdLeaveProcessService service;
        public AtdLeaveProcessController(ILogger<AtdLeaveProcessController> logger, IAtdLeaveProcessService service) : base(service)
        {
            this.logger = logger;
            this.service = service;
        }


        [HttpGet]
        [Authorize]
        [Route("GetUnapprovedUser")]
        public virtual JsonResultModel<bool> TestNotify()
        {
            service.GetUnapprovedUser();
            return new JsonResultModel<bool>
            {
                Code = "000",
                Message = "success",
                Result = true,
            };
        }

    }
}
