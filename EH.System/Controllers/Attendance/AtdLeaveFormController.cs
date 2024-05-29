using EH.System.Commons;
using EH.System.Models.Common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using EH.System.Models.Entities;
using EH.Service.Interface.Attendance;
using EH.System.Models.Dtos;
using static EH.Service.Implement.Attendance.AtdLeaveFormService;
using EH.Service.Implement;
using NPOI.SS.Formula.Functions;

namespace EH.System.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AtdLeaveFormController : BaseController<Atd_LeaveForm>
    {
        private readonly ILogger<AtdLeaveFormController> logger;
        private readonly IAtdLeaveFormService service;
        public AtdLeaveFormController(ILogger<AtdLeaveFormController> logger, IAtdLeaveFormService service) : base(service)
        {
            this.logger = logger;
            this.service = service;
        }

        [HttpGet]
        [Authorize]
        [Route("GetHomePageBodyData")]
        public virtual JsonResultModel<List<decimal>> GetHomePageBodyData(string userId)
        {
            var res = service.GetHomePageBodyData(userId);
            return new JsonResultModel<List<decimal>>
            {
                Code = "000",
                Message = "success",
                Result = res,
            };
        }

        [HttpGet]
        [Authorize]
        [Route("GetHomePageData")]
        public virtual JsonResultModel<Dictionary<string, int>> GetHomePageData(string userId)
        {
            var res = service.GetHomePageData(userId);
            return new JsonResultModel<Dictionary<string, int>>
            {
                Code = "000",
                Message = "success",
                Result = res,
            };
        }

        [HttpPost]
        [Authorize]
        [Route("GetStatic")]
        public virtual JsonResultModel<List<Atd_Statics>> GetStatic(PageRequest request)
        {
            var res = service.GetStatistic(request, out int total);
            return new JsonResultModel<List<Atd_Statics>>
            {
                Code = "000",
                Message = "success",
                Result = res,
                Other = total,
            };
        }

        [HttpPost]
        [Authorize]
        [Route("GetPageListWithState")]
        public virtual JsonResultModel<List<Atd_FormWithState>> GetPageListWithState([FromBody] PageRequest<Atd_LeaveForm> pageRequest)
        {
            var list = service.GetPageListWithState(pageRequest, out int total);
            return new JsonResultModel<List<Atd_FormWithState>>
            {
                Code = "000",
                Message = "success",
                Result = list,
                Other = total,
            };
        }



        [HttpPost]
        [Authorize]
        [Route("QueryGetPageList")]
        public virtual JsonResultModel<List<Atd_FormWithState>> GetPageList([FromBody] PageRequest<Atd_LeaveForm> pageRequest)
        {
            var list = service.QueryGetPageList(pageRequest, out int total);
            return new JsonResultModel<List<Atd_FormWithState>>
            {
                Code = "000",
                Message = "success",
                Result = list,
                Other = total,
            };
        }

        [HttpPost]
        [Authorize]
        [Route("Treate")]
        public virtual JsonResultModel<bool> Treate(List<string> ids)
        {
            var res = service.Treate(ids);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res,
            };
        }

        [HttpPost]
        [Authorize]
        [Route("GetWaitAuditForm")]
        public virtual JsonResultModel<List<Atd_FormAndProcess>> GetWaitAuditForm(PageRequest<Atd_LeaveForm> pageRequest)
        {
            var res = service.GetWaitAuditForm(pageRequest, out int totalCount);
            return new JsonResultModel<List<Atd_FormAndProcess>>
            {
                Code = "000",
                Message = "success",
                Result = res,
                Other = totalCount,
            };
        }

        public override JsonResultModel<Atd_LeaveForm> Add(Atd_LeaveForm entity, bool isSave = true)
        {
            return service.Apply(entity);
        }

        [HttpPost]
        [Authorize]
        [Route("Cancel")]
        public  JsonResultModel<bool> Cancel(Atd_LeaveForm entity)
        { 
            var res=service.Delete(entity);
            return new JsonResultModel<bool>
            {
                Code = res? "000":"100",
                Message= res ? "success" : "fail",
                Result= res,
            };
        }

        [HttpPost]
        [Authorize]
        [Route("UpdateFP")]
        public JsonResultModel<Atd_LeaveForm> UpdateFP(Atd_LeaveForm entity)
        {
            return service.UpdateFP(entity);
        }




        [HttpPost]
        [Authorize]
        [Route("Audit")]
        public virtual JsonResultModel<bool> Audit(Atd_Audit T)
        {
            var res = service.AuditForm(T);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "false",
                Result = res,
                Other = null,
            };
        }

        [HttpPost]
        [Authorize]
        [Route("Save")]
        public virtual JsonResultModel<Atd_LeaveForm> Save(Atd_LeaveForm T)
        {
            var res = service.Save(T);
            return new JsonResultModel<Atd_LeaveForm>
            {
                Code = res != null ? "000" : "100",
                Message = res != null ? "success" : "fail",
                Result = res
            };
        }


        [HttpPost]
        [Authorize]
        [Route("CalculateLeaveHours")]
        public virtual JsonResultModel<double> CalculateLeaveHours(DateTime dtStart, DateTime dtEnd, string workTime, bool isContainHoliday, bool isNursing)
        {
            var res = service.CalculateLeaveHours(dtStart, dtEnd, workTime,isContainHoliday, isNursing);
            return new JsonResultModel<double>
            {
                Code = res != null ? "000" : "100",
                Message = res != null ? "success" : "fail",
                Result = res
            };
        }

        

    }
}
