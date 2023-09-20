using EH.Service.Interface;
using EH.System.Models.Common;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EH.System.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class BaseController<T> : ControllerBase where T : class
    {

        private readonly IBaseService<T> baseService;
        public BaseController(IBaseService<T> baseService)
        {
            this.baseService = baseService;
        }


        [HttpPost]
        [Authorize]
        [Route("GetPageList")]
        public JsonResultModel<List<T>> GetPageList([FromBody] PageRequest<T> pageRequest)
        {
            var list = baseService.GetPageList(pageRequest, out int total);
            return new JsonResultModel<List<T>>
            {
                Code = "000",
                Message = "success",
                Result = list,
                Other = total,
            };
        }


        [HttpPost]
        [Authorize]
        [Route("Add")]
        public JsonResultModel<bool> Add(T entity)
        { 
            var res = baseService.Insert(entity);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }

        [HttpPost]
        [Authorize]
        [Route("Delete")]
        public JsonResultModel<bool> Delete(List<string> ids)
        {  
            var res = baseService.DeleteRange(ids);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }

        [HttpPost]
        [Authorize]
        [Route("RealDelete")]
        public JsonResultModel<bool> RealDelete(List<string> ids)
        {  
            var res = baseService.RealDeleteRange(ids);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }

        [HttpPost]
        [Authorize]
        [Route("Update")]
        public JsonResultModel<bool> Update(T entity)
        {
            var userName = HttpContext.User.Identity.Name.Split('\\')[1];
              
            var res = baseService.Update(entity);
            return new JsonResultModel<bool>
            {
                Code = res ? "000" : "100",
                Message = res ? "success" : "fail",
                Result = res
            };
        }
    }
}
