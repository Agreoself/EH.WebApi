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
    public interface IADUserOperateService : IBaseService<Sys_Users>
    {
        Sys_ADUsers GetUserByUserName(string userName);


    }
}
