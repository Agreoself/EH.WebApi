using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Interface.Attendance;
using EH.Service.Interface.Attendance;
using EH.Repository.Interface.Sys;
using System.Net.Sockets;
using System.Security.Cryptography.X509Certificates;
using NPOI.SS.Formula.Functions;

namespace EH.Service.Implement.Attendance
{
    public class AtdHolidaySettingService : BaseService<Atd_HolidaySetting>, IAtdHolidaySettingService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IAtdHolidaySettingRepository repository;
        public AtdHolidaySettingService(IAtdHolidaySettingRepository repository, LogHelper logHelper) : base(repository, logHelper)
        {
            this.repository = repository;
            this.logHelper = logHelper;
        }

        public List<string> GetHoliday(int year)
        {
            List<string> holidayStrings = new();
            DateTime startDate = new(year, 1, 1);
            DateTime endDate = new(year, 12, 31);

            var holidays = repository.Where(i => i.HolidayDate >= startDate && i.HolidayDate <= endDate).ToList();
            if (!holidays.Any())
            {
                List<Atd_HolidaySetting> atd_HolidaySettings = new();
                while (startDate <= endDate)
                {
                    if (startDate.DayOfWeek == DayOfWeek.Saturday || startDate.DayOfWeek == DayOfWeek.Sunday)
                    {
                        Atd_HolidaySetting atd_HolidaySetting = new Atd_HolidaySetting();
                        atd_HolidaySetting.HolidayDate = startDate;
                        atd_HolidaySetting.Remark = "First Create Weekend";
                        atd_HolidaySettings.Add(atd_HolidaySetting);
                        holidayStrings.Add(startDate.ToString("yyyy-MM-dd"));
                    }
                    startDate = startDate.AddDays(1);
                }
                repository.AddRange(atd_HolidaySettings);
                return holidayStrings;
            }
            else
            {
                return holidays.Select(i => i.HolidayDate.ToString("yyyy-MM-dd")).ToList();
            }
        }

        public bool SaveHoliday(List<string> holidays)
        {
            if (holidays.Any())
            {
                List<Atd_HolidaySetting> saveHolidays = new();
                foreach (var item in holidays)
                {
                    Atd_HolidaySetting atd_HolidaySetting = new Atd_HolidaySetting();
                    atd_HolidaySetting.HolidayDate = Convert.ToDateTime(item);
                    saveHolidays.Add(atd_HolidaySetting);
                }

                var year = Convert.ToDateTime(holidays[0]).Year;
                DateTime startDate = new(year, 1, 1);
                DateTime endDate = new(year, 12, 31);
                var holidayList = repository.Where(i => i.HolidayDate >= startDate && i.HolidayDate <= endDate).ToList();

                try
                {
                    repository.DeleteRange(holidayList);
                    repository.AddRange(saveHolidays);
                    return true;
                }
                catch (Exception ex)
                {
                    logHelper.LogError("SaveHoliday Error" + ex.ToString());
                    return false;
                }

            }
            else
            { 
                return false;
            }
        }

    }
}
