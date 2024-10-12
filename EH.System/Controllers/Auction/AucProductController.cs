using Microsoft.AspNetCore.Mvc;
using EH.Service.Interface.Sys;
using EH.System.Models.Entities;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using NPOI.SS.Formula.Functions;
using System.Collections.Generic;

namespace EH.System.Controllers.Auction
{
    [Route("api/[controller]")]
    [ApiController]
    public class AucProductController : BaseController<Auc_Product>
    {
        private readonly ILogger<AucProductController> logger;
        private readonly IAucProductService productService;
        public AucProductController(ILogger<AucProductController> logger, IAucProductService productService) : base(productService)
        {
            this.logger = logger;
            this.productService = productService;
        }

        [HttpPost("UploadIamges")]
        public JsonResultModel<string> Upload([FromForm] string product, [FromForm] string? originImgs)
        {
            var files = Request.Form.Files;

            var res = productService.UploadAttachments(files, product, originImgs).Result;

            return new JsonResultModel<string>
            {
                Code = res == "Exception" ? "100" : "000",
                Result = res,
                Message = res == "Exception" ? "Fail" : "Success"
            };
        }

        [HttpPost("DiyInsert")]
        public JsonResultModel<Auc_Product> DiyInsert([FromForm] Auc_ProductDto dto)
        {

            var res = productService.DiyInsert(dto);

            return new JsonResultModel<Auc_Product>
            {
                Code = res == null ? "100" : "000",
                Result = res,
                Message = res == null ? "Fail" : "Success"
            };
        }

        [HttpPost("Import")]
        public JsonResultModel<bool> Import([FromForm] string activityId)
        {
            var files = Request.Form.Files;

            if (files.Count > 0)
            {
                var res = productService.Import(files[0], activityId);
                return new JsonResultModel<bool>
                {
                    Code = res ? "000" : "100",
                    Result = res,
                    Message = res ? "Success" : "Fail"
                };
            }
            else
            {
                return new JsonResultModel<bool>
                {
                    Code = "100",
                    Result = false,
                    Message = "Fail"
                };
            }
        }

        [HttpPost("Export")]
        public JsonResultModel<string> Export(string activityId)
        {
            var res = productService.Export(activityId);
            string base64String = Convert.ToBase64String(res);
            return new JsonResultModel<string>
            {
                Code = res!=null?"000":"100",
                Result = base64String,
                Message = res != null?"Success":"Fail"
            };
        }

        [HttpPost("Search")]
        public JsonResultModel<List<Auc_Product>> Search(Auc_ProductSeachDto dto)
        {
            var res = productService.Search(dto,out int total);
            return new JsonResultModel<List<Auc_Product>>
            {
                Code = "000",
                Message = "success",
                Result = res,
                Other = total,
            };  
        }

        [HttpPost("GetMyAuctions")]
        public JsonResultModel<List<Auc_MyAuctionDto>> GetMyAuctions(Auc_MyAuctionRequestDto dto)
        {
            var res = productService.GetMyAuctions(dto, out int total);
            return new JsonResultModel<List<Auc_MyAuctionDto>>
            {
                Code = "000",
                Message = "success",
                Result = res,
                Other = total,
            };
        }
    }
}
