using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Commons
{ 
    public static class DateTimeHelper
    {
        /// <summary>
        /// 时间戳转本地时间-时间戳精确到秒
        /// </summary>
        /// <param name="seconds"></param>
        /// <returns></returns>
        public static DateTime FromUnixTimeSeconds(long seconds)
        {
            var dto = DateTimeOffset.FromUnixTimeSeconds(seconds);
            return dto.ToLocalTime().DateTime;
        }

        /// <summary>
        /// 时间戳转本地时间-时间戳精确到毫秒
        /// </summary>
        /// <returns></returns>
        public static DateTime FromUnixTimeMilliseconds(long milliseconds)
        {

            var dto = DateTimeOffset.FromUnixTimeMilliseconds(milliseconds);
            return dto.ToLocalTime().DateTime;
        }

        /// <summary>
        /// 时间转时间戳Unix-时间戳精确到秒
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static long ToUnixTimeSeconds(DateTime? dateTime = null)
        {
            if (!dateTime.HasValue) dateTime = DateTime.Now;
            DateTimeOffset dto = new DateTimeOffset(dateTime.Value);
            return dto.ToUnixTimeSeconds();
        }

        /// <summary>
        /// 时间转时间戳Unix-时间戳精确到毫秒
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public static long ToUnixTimeMilliseconds(DateTime? dateTime = null)
        {
            if (!dateTime.HasValue) dateTime = DateTime.Now;
            DateTimeOffset dto = new DateTimeOffset(dateTime.Value);
            return dto.ToUnixTimeMilliseconds();
        }


        //public static double CalculateTotalHours(DateTime dtStart, DateTime dtEnd,Dictionary<string,object> workDays=null)
        //{

        //    DateTime dtFirstDayGoToWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 8, 30, 0);//请假第一天的上班时间
        //    DateTime dtFirstDayGoOffWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 17, 30, 0);//请假第一天的下班时间

        //    DateTime dtLastDayGoToWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 8, 30, 0);//请假最后一天的上班时间
        //    DateTime dtLastDayGoOffWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 17, 30, 0);//请假最后一天的下班时间

        //    DateTime dtFirstDayRestStart = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 12, 00, 0);//请假第一天的午休开始时间
        //    DateTime dtFirstDayRestEnd = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 13, 00, 0);//请假第一天的午休结束时间

        //    DateTime dtLastDayRestStart = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 12, 00, 0);//请假最后一天的午休开始时间
        //    DateTime dtLastDayRestEnd = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 13, 00, 0);//请假最后一天的午休结束时间

        //    //如果开始请假时间早于上班时间或者结束请假时间晚于下班时间，者需要重置时间
        //    if (!IsWorkDay(dtStart, workDays) && !IsWorkDay(dtEnd, workDays))
        //        return 0;
        //    if (dtStart >= dtFirstDayGoOffWork && dtEnd <= dtLastDayGoToWork && (dtEnd - dtStart).TotalDays < 1)
        //        return 0;
        //    if (dtStart >= dtFirstDayGoOffWork && !IsWorkDay(dtEnd, workDays) && (dtEnd - dtStart).TotalDays < 1)
        //        return 0;

        //    if (dtStart < dtFirstDayGoToWork)//早于上班时间
        //        dtStart = dtFirstDayGoToWork;
        //    if (dtStart >= dtFirstDayGoOffWork)//晚于下班时间
        //    {
        //        while (dtStart < dtEnd)
        //        {
        //            dtStart = new DateTime(dtStart.AddDays(1).Year, dtStart.AddDays(1).Month, dtStart.AddDays(1).Day, 8, 30, 0);
        //            if (IsWorkDay(dtStart, workDays))
        //            {
        //                dtFirstDayGoToWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 8, 30, 0);//请假第一天的上班时间
        //                dtFirstDayGoOffWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 17, 30, 0);//请假第一天的下班时间
        //                dtFirstDayRestStart = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 11, 30, 0);//请假第一天的午休开始时间
        //                dtFirstDayRestEnd = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 12, 30, 0);//请假第一天的午休结束时间

        //                break;
        //            }
        //        }
        //    }

        //    if (dtEnd > dtLastDayGoOffWork)//晚于下班时间
        //        dtEnd = dtLastDayGoOffWork;
        //    if (dtEnd <= dtLastDayGoToWork)//早于上班时间
        //    {
        //        while (dtEnd > dtStart)
        //        {
        //            dtEnd = new DateTime(dtEnd.AddDays(-1).Year, dtEnd.AddDays(-1).Month, dtEnd.AddDays(-1).Day, 17, 30, 0);
        //            if (IsWorkDay(dtEnd, workDays))//
        //            {
        //                dtLastDayGoToWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 8, 30, 0);//请假最后一天的上班时间
        //                dtLastDayGoOffWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 17, 30, 0);//请假最后一天的下班时间
        //                dtLastDayRestStart = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 11, 30, 0);//请假最后一天的午休开始时间
        //                dtLastDayRestEnd = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 12, 30, 0);//请假最后一天的午休结束时间
        //                break;
        //            }
        //        }
        //    }

        //    //计算请假第一天和最后一天的小时合计数并换算成分钟数           
        //    double iSumMinute = dtFirstDayGoOffWork.Subtract(dtStart).TotalMinutes + dtEnd.Subtract(dtLastDayGoToWork).TotalMinutes;//计算获得剩余的分钟数

        //    if (dtStart > dtFirstDayRestStart && dtStart < dtFirstDayRestEnd)
        //    {//开始休假时间正好是在午休时间内的，需要扣除掉
        //        iSumMinute = iSumMinute - dtFirstDayRestEnd.Subtract(dtStart).Minutes;
        //    }
        //    if (dtStart < dtFirstDayRestStart)
        //    {//如果是在午休前开始休假的就自动减去午休的60分钟
        //        iSumMinute = iSumMinute - 60;
        //    }
        //    if (dtEnd > dtLastDayRestStart && dtEnd < dtLastDayRestEnd)
        //    {//如果结束休假是在午休时间内的，例如“请假截止日是1月31日 12:00分”的话那休假时间计算只到 11:30分为止。
        //        iSumMinute = iSumMinute - dtEnd.Subtract(dtLastDayRestStart).Minutes;
        //    }
        //    if (dtEnd > dtLastDayRestEnd)
        //    {//如果是在午休后结束请假的就自动减去午休的60分钟
        //        iSumMinute = iSumMinute - 60;
        //    }


        //    int leaveday = 0;//实际请假的天数
        //    double countday = 0;//获取两个日期间的总天数

        //    DateTime tempDate = dtStart;//临时参数
        //    while (tempDate < dtEnd)
        //    {
        //        countday++;
        //        tempDate = new DateTime(tempDate.AddDays(1).Year, tempDate.AddDays(1).Month, tempDate.AddDays(1).Day, 0, 0, 0);
        //    }
        //    //循环用来扣除双休日、法定假日 和 添加调休上班
        //    for (int i = 0; i < countday; i++)
        //    {
        //        DateTime tempdt = dtStart.Date.AddDays(i);

        //        if (IsWorkDay(tempdt, workDays))
        //            leaveday++;
        //    }

        //    //去掉请假第一天和请假的最后一天，其余时间全部已8小时计算。
        //    //SumMinute/60： 独立计算 请假第一天和请假最后一天总归请了多少小时的假
        //    double doubleSumHours = RoundToNearestHalf(((leaveday - 2) * 8) + iSumMinute / 60);
        //    //int intSumHours = Convert.ToInt32(doubleSumHours);

        //    //if (doubleSumHours > intSumHours)//如果请假时间不足1小时话自动算作1小时
        //    //    intSumHours++;

        //    return doubleSumHours;

        //}

        //private static bool IsWorkDay(DateTime date,Dictionary<string,object> workDays)
        //{
        //    try
        //    {
        //        //读取数据库中【Vacation】表中的所有数据,返回一个Datatable等等...
        //        //我这里采用的是内存操作Dictionary，因为一般这种节假日都是固定不变的，不需要每次都取访问数据查询一遍。

        //        //Dictionary<string, Vacation> m_simVacationData = newDictionary<string, SimVacation>(); 

        //        //利用Datatable的值循环给 m_simVacationData 赋值。

        //    }
        //    catch (Exception)
        //    {
        //        //抛出异常
        //    }

        //    string DateKey = date.ToString("yyyy-MM-dd");//日期值：“2012-08-01”

        //    bool b_wokrdate = true;


        //    ////星期天并且不属于节假日和调休上班
        //    //if (date.DayOfWeek == DayOfWeek.Sunday && !m_simVacationData.ContainsKey(DateKey))
        //    //    return false;
        //    ////星期六并且不属于节假日和调休上班     

        //    //else if (date.DayOfWeek == DayOfWeek.Saturday && !m_simVacationData.ContainsKey(DateKey))

        //    //    return false;
        //    //else if (m_simVacationData.ContainsKey(DateKey))//属于节假日或调休
        //    //{
        //    //    if (m_simVacationData[DateKey].Is_Vacation)//Is_Vacation=true（节假日） Is_Vacation=false（调休上班）
        //    //    {
        //    //        return false;
        //    //    }
        //    //}
        //    if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
        //    {
        //        b_wokrdate = false;
        //    }

        //    return b_wokrdate;
        //}

        public static double RoundToNearestHalf(double value)
        {
            int rounded = (int)Math.Round(value / 0.5);
            return rounded * 0.5;
        }
    }
}
