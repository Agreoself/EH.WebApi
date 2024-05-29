using EH.System.Commons;
using EH.System.Models.Common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using EH.Service.Interface.Sys;
using EH.Service.Interface.AD;
using EH.System.Models.Entities;

namespace EH.System.Controllers.AD
{
    [Route("api/[controller]")]
    [ApiController]
    public class ADUserPwdNotifyController : ControllerBase
    {
        private readonly ILogger<ADUserPwdNotifyController> logger;
        private readonly IADUserPwdNotifyService adService;
        public ADUserPwdNotifyController(ILogger<ADUserPwdNotifyController> logger, IADUserPwdNotifyService adService)
        {
            this.logger = logger;
            this.adService = adService;
        }


        ///// <summary>
        ///// 重置AD用户的密码
        ///// </summary>
        ///// <param name="userName">AD用户名</param>
        ///// <param name="isContinuous">是否连续</param>
        ///// <param name="startIndex">开始索引</param>
        ///// <param name="limit">多少用户</param>
        //[Authorize]
        //[HttpGet]
        //[Route("restpassword")]
        //public void ADResetPassword(string userName, bool isContinuous, int startIndex, int limit)
        //{
        //    string realUserName = userName;
        //    if (isContinuous)
        //    {
        //        Dictionary<string, string> dict = new Dictionary<string, string>();
        //        int length = Math.Abs(limit).ToString().Length;
        //        for (int i = startIndex; i <= limit; i++)
        //        {
        //            realUserName = userName + i.ToString().PadLeft(length, '0');
        //            string pwd = adService.ResetPassword(realUserName).Result;
        //            logger.LogInformation(realUserName + ":" + pwd);
        //            dict.Add(realUserName, pwd);
        //        }

        //        OfficeHelper.writeToExcel(dict);
        //    }
        //    else
        //    {
        //        string pwd = adService.ResetPassword(realUserName).Result;
        //        logger.LogInformation(realUserName + ":" + pwd);
        //    }
        //}


        [HttpGet]
        [Authorize]
        [Route("Get")]
        public JsonResultModel<List<AD_UserPwdNotify>> GetADUser(string location)
        {
            var res = adService.GenerateADUser(location);
            return new JsonResultModel<List<AD_UserPwdNotify>>
            {
                Code = res!=null ? "000" : "100",
                Result = res ,
                Message = res != null ? "Success" : "Fail",
            };
        }

        [HttpGet]
        [Authorize]
        [Route("SendNotify")]
        public virtual bool TestEmil(string userId)
        {
            var res = adService.SendMail(userId);
            return res;
        }


        [HttpGet]
        [Authorize]
        [Route("Test")]
        public virtual string Test(decimal? totalHours)
        {
            var res = adService.Test(totalHours);
            return res;
        }


        [HttpGet]
        [Authorize]
        [Route("TestSendMail")]
        public virtual string TestSendMail(string fromEmail, string toEmail, string title, string body)
        {
            var res = adService.TestSendMail(fromEmail,toEmail,title,body);
            return res;
        }





    }
}
