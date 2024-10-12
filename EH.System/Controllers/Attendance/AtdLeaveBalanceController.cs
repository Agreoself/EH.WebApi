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
    public class AtdLeaveBalanceController : BaseController<Atd_LeaveBalance>
    {
        private readonly ILogger<AtdLeaveBalanceController> logger;
        private readonly IAtdLeaveBalanceService service;
        public AtdLeaveBalanceController(ILogger<AtdLeaveBalanceController> logger, IAtdLeaveBalanceService service) : base(service)
        {
            this.logger = logger;
            this.service = service;
        }

        [HttpGet]
        [Authorize]
        [Route("testcacule")]
        public virtual double TestEmil(DateTime workDate, DateTime ehcDate)
        {
            var s = new LeaveBalanceHelper(workDate, ehcDate)._sickTotalHour;
            return s;
        }

        [HttpPost]
        [Authorize]
        [Route("CalculateAnnualAndSick")]
        public virtual JsonResultModel<bool> CalculateAnnualAndSick(List<string> userIds)
        {
            var res = service.CalculateAnnualAndSick(userIds);
            return new JsonResultModel<bool>
            {
                Result = res,
                Code = res ? "000" : "100",
                Message =res? "success":"false",
            };
        }

        [HttpPost]
        [Authorize]
        [Route("CalculatePersonal")]
        public virtual JsonResultModel<bool> CalculatePersonal(List<string> userIds)
        {
            var res = service.CalculatePersonal(userIds);
            return new JsonResultModel<bool>
            {
                Result = res,
                Code = res ? "000" : "100",
                Message = res ? "success" : "false",
            };
        }

        [HttpPost]
        [Authorize]
        [Route("GetInfo")]
        public virtual JsonResultModel<List<Atd_AnnualInfos>> GetInfo(Atd_AnnualInfoReq data)
        {
            var res = service.GetInfo(data);
            return new JsonResultModel<List<Atd_AnnualInfos>>
            {
                Result = res,
                Code = res!=null ? "000" : "100",
                Message = res != null ? "success" : "false",
            };
        }

        [HttpPost]
        [Authorize]
        [Route("Statistics")]
        public virtual JsonResultModel<List<Atd_BalanceStatics>> Statistics(PageRequest request)
        {
            var res = service.Statistics(request,out int totalCount);
            return new JsonResultModel<List<Atd_BalanceStatics>>
            {
                Result = res,
                Code = res != null ? "000" : "100",
                Message = res != null ? "success" : "false",
                Other = totalCount
            };
        }

        [HttpGet]
        [Authorize]
        [Route("ClearCarryoverAnnual")]
        public virtual JsonResultModel<string> ClearCarryoverAnnual(bool isClear)
        {
            var res = service.ClearCarryoverAnnual(isClear);
            return new JsonResultModel<string>
            {
                Result = res,
                Code = res!="false" ? "000" : "100",
                Message = res != "false" ? "success" : "false",
            };
        }
    }

}
