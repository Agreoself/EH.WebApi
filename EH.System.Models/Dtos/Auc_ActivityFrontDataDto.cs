using EH.System.Models.Common;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Auc_ActivityFrontDataDto : Auc_Activity,NoEntity
    {
        public string Time => StartTime.ToString("yyyy-MM-dd HH:mm:ss") +" - "+ EndTime.ToString("yyyy-MM-dd HH:mm:ss");

        public string Status => Lifecycle == 0 ? "info" : Lifecycle == 1 ? "success" : "danger";
        public string StatusText => Lifecycle == 0 ? "即将开始" : Lifecycle == 1 ? "正在进行" : "已结束";


    }
}
