using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.AspNetCore.Mvc;

namespace EH.System.Attribute
{
    public class CustomExceptionFilterAttribute
    { 
    }
    //<T> where T : ExceptionFilterAttribute,
    //{
    //    private readonly ILogger<T> _logger;

    //    public CustomExceptionFilterAttribute(ILogger<T> logger)
    //    {
    //        _logger = logger;
    //    }

    //    public override void OnException(ExceptionContext context)
    //    {
    //        //判断该异常有没有处理
    //        if (!context.ExceptionHandled)
    //        {
    //            _logger.LogError($"Path:{context.HttpContext.Request.Path}Message:{context.Exception.Message}");
    //            context.Result = new JsonResult(new
    //            {
    //                Reslut = false,
    //                Msg = "发生异常，请联系管理员"
    //            });
    //            context.ExceptionHandled = true;
    //        }
    //    }
    //}

}
