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

        [HttpGet]
        [Authorize]
        [Route("GetEntityById")]
        public virtual JsonResultModel<T> GetEntityById(Guid id)
        {
            var entity = baseService.GetEntityById(id);
            return new JsonResultModel<T>
            {
                Code = entity!=null?"000":"100",
                Message = entity != null ? "success":"fail",
                Result = entity,
            };
        }

        [HttpPost]
        [Authorize]
        [Route("GetPageList")]
        public virtual JsonResultModel<List<T>> GetPageList([FromBody] PageRequest<T> pageRequest)
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
        public virtual JsonResultModel<T> Add(T entity, bool isSave=true)
        { 
            var res = baseService.Insert(entity, isSave);
            return new JsonResultModel<T>
            {
                Code = res!=null ? "000" : "100",
                Message = res != null ? "success" : "fail",
                Result = res
            };
        }

        [HttpPost]
        [Authorize]
        [Route("Delete")]
        public virtual JsonResultModel<bool> Delete(List<string> ids)
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
        public virtual JsonResultModel<bool> RealDelete(List<string> ids)
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
        public virtual JsonResultModel<bool> Update(T entity)
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
