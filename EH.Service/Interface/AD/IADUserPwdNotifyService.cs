using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.AD
{
    public interface IADUserPwdNotifyService : IBaseService<AD_UserPwdNotify>
    {

        List<AD_UserPwdNotify> GenerateADUser(string location);
        bool SendMail(string userId);
        void SendMail(List<AD_UserPwdNotify> users);

        string Test(decimal? userId);
        public string TestSendMail(string fromEmail, string toEmail, string title, string body);
    }
}
