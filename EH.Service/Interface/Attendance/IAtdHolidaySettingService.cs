using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.Attendance
{
    public interface IAtdHolidaySettingService : IBaseService<Atd_HolidaySetting>
    {
        List<string> GetHoliday(int year);

        bool SaveHoliday(List<string> holidays);
    }
}
