using EH.System.Commons;
using EH.System.Models.Common;
using EH.Service.Interface;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;


namespace EH.System.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ADUserController : ControllerBase
    {
        private readonly ILogger<ADUserController> logger;
        private readonly IADUserService adService;
        public ADUserController(ILogger<ADUserController> logger, IADUserService adService)
        {
            this.logger = logger;
            this.adService = adService;
        }


        /// <summary>
        /// 重置AD用户的密码
        /// </summary>
        /// <param name="userName">AD用户名</param>
        /// <param name="isContinuous">是否连续</param>
        /// <param name="startIndex">开始索引</param>
        /// <param name="limit">多少用户</param>
        [Authorize]
        [HttpGet]
        [Route("restpassword")]
        public void ADResetPassword(string userName, bool isContinuous, int startIndex, int limit)
        {
            string realUserName = userName;
            if (isContinuous)
            {
                Dictionary<string, string> dict = new Dictionary<string, string>();
                int length = Math.Abs(limit).ToString().Length;
                for (int i = startIndex; i <= limit; i++)
                {
                    realUserName = userName + i.ToString().PadLeft(length, '0');
                    string pwd = adService.ResetPassword(realUserName).Result;
                    logger.LogInformation(realUserName + ":" + pwd);
                    dict.Add(realUserName, pwd);
                }

                OfficeHelper.writeToExcel(dict);
            }
            else
            {
                string pwd = adService.ResetPassword(realUserName).Result;
                logger.LogInformation(realUserName + ":" + pwd);
            }
        }


        /// <summary>
        /// 获取AD用户信息
        /// </summary>
        /// <param name="userName">AD用户名</param>
        [Authorize]
        [HttpGet]
        [Route("getaduser")]
        public JsonResultModel<object> GetADUser(string userName)
        {
            var res = adService.GetUser(userName);  
            return res;
        }

        [HttpGet]
        [Route("ADLogin")]
        public JsonResultModel<bool> ADLogin(string userName, string password)
        {
            return adService.CheckLogin(userName,password);
        }

        [HttpGet]
        [Authorize]
        [Route("GenerateUser")]
        public JsonResultModel<bool> GenerateUser()
        {
            return adService.GenerateUser();
        }

    }
}
