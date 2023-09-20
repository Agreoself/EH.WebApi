

namespace EH.System.Attribute
{
    using EH.System.Commons;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.Mvc.Filters;  

    public class TokenValidationAttribute : ActionFilterAttribute
    {
        private readonly JwtHelper jwtHelper;

        public TokenValidationAttribute(JwtHelper jwtHelper)
        {
            this.jwtHelper = jwtHelper;
        }

        public override async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
        {
            if (!context.HttpContext.Request.Headers.ContainsKey("Authorization"))
            {
                context.Result = new UnauthorizedResult();
                return;
            }

            var token = context.HttpContext.Request.Headers["Authorization"];
            if (!jwtHelper.ValidateToken(token))
            {
                context.Result = new UnauthorizedResult();
                return;
            }
            else
            {
                await next.Invoke();
            }
        }
    }
}
