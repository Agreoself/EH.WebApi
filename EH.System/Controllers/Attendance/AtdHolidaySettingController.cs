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
    public class AtdHolidaySettingController : BaseController<Atd_HolidaySetting>
    {
        private readonly ILogger<AtdHolidaySettingController> logger;
        private readonly IAtdHolidaySettingService service;
        public AtdHolidaySettingController(ILogger<AtdHolidaySettingController> logger, IAtdHolidaySettingService service) : base(service)
        {
            this.logger = logger;
            this.service = service;
        }

        [HttpGet]
        [Authorize]
        [Route("GetHoliday")]
        public virtual JsonResultModel<List<string>> GetHoliday(int year)
        {
            var res =service.GetHoliday(year);
            return new JsonResultModel<List<string>>
            {
                Result = res,
                Code = "000",
                Message = "success",
            };
        }

        [HttpPost]
        [Authorize]
        [Route("SaveHoliday")]
        public virtual JsonResultModel<bool> SaveHoliday(List<string> holidays)
        {
            var res = service.SaveHoliday(holidays);
            return new JsonResultModel<bool>
            {
                Result = res,
                Code = res?"000":"100",
                Message = res?"Success":"Fail",
            };
        }
    }

}
