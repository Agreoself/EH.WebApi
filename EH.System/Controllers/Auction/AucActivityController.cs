using Microsoft.AspNetCore.Mvc;
using EH.Service.Interface.Sys;
using EH.System.Models.Entities;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Commons;
using Microsoft.AspNetCore.Authorization;

namespace EH.System.Controllers.Auction
{
    [Route("api/[controller]")]
    [ApiController]
    public class AucActivityController : BaseController<Auc_Activity>
    {
        private readonly ILogger<AucActivityController> logger;
        private readonly IAucActivityService activityService;
        public AucActivityController(ILogger<AucActivityController> logger, IAucActivityService activityService):base(activityService)
        {
            this.logger = logger;
            this.activityService = activityService;
        }


        [HttpPost]
        [Authorize]
        [Route("FrontGetPageList")]
        public JsonResultModel<List<Auc_ActivityFrontDataDto>> FrontGetPageList([FromBody] PageRequest<Auc_Activity> pageRequest)
        {
            var s = base.GetPageList(pageRequest);
            return s.ToObject<JsonResultModel<List<Auc_ActivityFrontDataDto>>>();
        }

    }
}
