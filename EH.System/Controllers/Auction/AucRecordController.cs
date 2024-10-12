using Microsoft.AspNetCore.Mvc;
using EH.Service.Interface.Sys;
using EH.System.Models.Entities;
using EH.System.Models.Common;
using Microsoft.AspNetCore.SignalR;
using System.DirectoryServices.Protocols;
using Microsoft.AspNetCore.Authorization;
using Microsoft.EntityFrameworkCore;

namespace EH.System.Controllers.Auction
{
    [Route("api/[controller]")]
    [ApiController]
    public class AucRecordController : BaseController<Auc_Record>
    {
        private readonly ILogger<AucRecordController> logger;
        private readonly IAucRecordService recordService; 
        public AucRecordController(ILogger<AucRecordController> logger, IAucRecordService recordService) :base(recordService)
        {
            this.logger = logger;
            this.recordService = recordService;
            //this.hubContext = hubContext; 
        }

        [HttpPost("Export")]
        public JsonResultModel<string> Export(string productId)
        {
            var res = recordService.Export(productId);
            string base64String = Convert.ToBase64String(res);
            return new JsonResultModel<string>
            {
                Code = res != null ? "000" : "100",
                Result = base64String,
                Message = res != null ? "Success" : "Fail"
            };


        }

        [HttpPost]
        [Route("placeBid")]
        [Authorize]
        public async Task<JsonResultModel<bool>> PlaceBid(Auc_Record entity)
        { 
            var res= await recordService.PlaceBidAsync(entity);
          
            //if (res) {
            //    //await hubContext.Clients.All.SendAsync("ReceiveBid", entity);
            //    await myAuctionsHub.MyAuctionReceiveBid(entity);
            //    await bidHub.SendBid(entity);
            //}
            // 广播出价消息
            return new JsonResultModel<bool>
            {
                Code = res  ? "000" : "100",
                Result = res,
                Message = res ? "Success" : "Fail"
            }; ;
        }

        [HttpGet("getCurrentPrice")]
        public JsonResultModel<decimal> GetCurrentPrice(string productId)
        {
            var res = recordService.GetCurrentPrice(productId); 
            return new JsonResultModel<decimal>
            {
                Code = res != 0 ? "000" : "100",
                Result = res,
                Message = res != 0 ? "Success" : "Fail"
            };


        }
    }
}
