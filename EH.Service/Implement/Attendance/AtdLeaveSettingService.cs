using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Interface.Attendance;
using EH.Service.Interface.Attendance;
using EH.Repository.DataAccess;
using EH.System.Models.Dtos;

namespace EH.Service.Implement.Attendance
{
    public class AtdLeaveSettingService : BaseService<Atd_LeaveSetting>, IAtdLeaveSettingService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IAtdLeaveSettingRepository repository;
        private readonly IAtdLeaveBalanceRepository balanceRepository;
        private readonly IAtdLeaveFormRepository formRepository;
        private readonly IAtdHolidaySettingRepository holidayRepository;
        public AtdLeaveSettingService(IAtdLeaveSettingRepository repository, IAtdLeaveBalanceRepository balanceRepository, IAtdLeaveFormRepository formRepository, IAtdHolidaySettingRepository holidayRepository, LogHelper logHelper) : base(repository, logHelper)
        {
            this.logHelper = logHelper;
            this.repository = repository;
            this.balanceRepository = balanceRepository;
            this.formRepository = formRepository;
            this.holidayRepository = holidayRepository;
        }

        public LeaveDetail GetLeaveDetail(string leaveType, string userId)
        {
            switch (leaveType)
            {
                case "annual": return GetBalanceDetail(userId, leaveType);
                case "sick": return GetBalanceDetail(userId, leaveType);
                case "personal": return GetBalanceDetail(userId, leaveType);
                case "parental": return GetBalanceDetail(userId, leaveType);
                default: return GetOtherDetail(leaveType);
            }
        }

        public Dictionary<int, List<string>> GetHoliday()
        {
            Dictionary<int, List<string>> holidays = new Dictionary<int, List<string>>();
            var currentYear = DateTime.Now.Year;
            var nextYear = currentYear + 1;
            var nowHolidays = holidayRepository.Where(i => i.HolidayDate.Year == currentYear).Select(o => o.HolidayDate.ToString("yyyy-MM-dd")).ToList();
            var nextHolidays = holidayRepository.Where(i => i.HolidayDate.Year == nextYear).Select(o => o.HolidayDate.ToString("yyyy-MM-dd")).ToList();
            holidays.Add(currentYear, nowHolidays);
            holidays.Add(nextYear, nextHolidays);
            return holidays;
        }
        public LeaveDetail GetBalanceDetail(string userId, string leaveType)
        {
            var year = DateTime.Now.Year;
            var balanceEntity = balanceRepository.FirstOrDefault(x => x.UserId == userId && x.LeaveType == leaveType&&x.Year==year);
            var settingEntity = repository.FirstOrDefault(i => i.LeaveType == leaveType);
            if (balanceEntity == null || settingEntity == null)
            {
                return new LeaveDetail
                {
                    Code = -1,
                    LeaveType = leaveType,
                    Available = 0,
                    IsHoliday = false,
                    needHr=true,
                    MinUnit = "All",
                };
            }
            else
            {
                return new LeaveDetail
                {
                    Code = 0,
                    LeaveType = leaveType,
                    Available = balanceEntity.Available,
                    IsHoliday = settingEntity.IsContainHoliday,
                    MinUnit = settingEntity.MinUnit,
                    needHr = settingEntity.NeedHRApprove,
                    holidays = GetHoliday(),
                };
            }
        }
   
        public LeaveDetail GetOtherDetail(string leaveType)
        {
            var settingEntity = repository.FirstOrDefault(i => i.LeaveType == leaveType);
            if (settingEntity == null)
            {
                return new LeaveDetail
                {
                    Code = -1,
                    LeaveType = leaveType,
                    Available = 0,
                    IsHoliday = false,
                    needHr =true, 
                    MinUnit = "All",
                };
            }
            else
            {
                return new LeaveDetail
                {
                    Code = 0,
                    LeaveType = leaveType,
                    Available = settingEntity.Limit,
                    IsHoliday = settingEntity.IsContainHoliday,
                    MinUnit = settingEntity.MinUnit,
                    needHr = settingEntity.NeedHRApprove,
                    holidays = GetHoliday(),
                };
            }
        }
    }
}
