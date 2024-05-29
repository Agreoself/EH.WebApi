using EH.System.Commons;
using EH.System.Models.Common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using EH.System.Models.Entities;
using EH.Service.Interface.Attendance;
using EH.System.Models.Dtos;

namespace EH.System.Controllers.Attendance
{
    [Route("api/[controller]")]
    [ApiController]
    public class AtdLeaveSettingController : BaseController<Atd_LeaveSetting>
    {
        private readonly ILogger<AtdLeaveSettingController> logger;
        private readonly IAtdLeaveSettingService service;
        public AtdLeaveSettingController(ILogger<AtdLeaveSettingController> logger, IAtdLeaveSettingService service) : base(service)
        {
            this.logger = logger;
            this.service = service;
        }

        [HttpPost]
        [Authorize]
        [Route("GetLeaveDetail")]
        public JsonResultModel<LeaveDetail> GetLeaveDetail(LeaveSettingRequest request)
        {
            var res = service.GetLeaveDetail(request.leaveType, request.userId);
            return new JsonResultModel<LeaveDetail>
            {
                Code=res.Code>=0?"000":"100",
                Message=res.Code>=0?"success":"fail",
                Result = res
            };
        }

    }
}
