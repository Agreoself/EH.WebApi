using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface
{
    public interface ISysUserService:IBaseService<Sys_Users>
    {
        public bool Login(User user);

        public UserInfo GetUserInfo(string userID);

        public bool GrantRole(List<string> userIds,List<string> roleIds);

        List<Sys_Users> GetUserListInRole(PageRequest<Sys_Users> request,out int totalCount);

    }
}
