using EH.System.Commons;
using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface
{
    public interface IADUserService
    {
        public JsonResultModel<object> GetUser(string userName);
        public JsonResultModel<string> ResetPassword(string userName);
        public JsonResultModel<bool> CheckLogin(string userName,string password);
        public JsonResultModel<bool> GenerateUser();
    }
}
