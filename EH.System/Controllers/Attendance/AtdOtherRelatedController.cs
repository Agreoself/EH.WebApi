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
    public class AtdOtherRelatedController : BaseController<Atd_OtherRelated>
    {
        private readonly ILogger<AtdOtherRelatedController> logger;
        private readonly IAtdOtherRelatedService service;
        public AtdOtherRelatedController(ILogger<AtdOtherRelatedController> logger, IAtdOtherRelatedService service) : base(service)
        {
            this.logger = logger;
            this.service = service;
        }
 
    }

}
