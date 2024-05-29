using EH.System.Models.Dtos;
using NPOI.SS.Formula.Functions;
using NPOI.SS.Formula.PTG;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using static EH.System.Commons.LeaveBalanceHelper;

namespace EH.System.Commons
{
    public class LeaveBalanceHelper
    {
        private readonly DateTime _workStartDate;
        private readonly DateTime _ehcStartDate;
        private readonly DateTime _currentEhcDate;
        private readonly DateTime _currentWorkDate;
        private readonly DateTime _now = DateTime.Now;
        private readonly DateTime _startDate = new DateTime(DateTime.Now.Year, 1, 1);
        private readonly DateTime _endDate = new DateTime(DateTime.Now.Year, 12, 31);
        public readonly double _annualTotalHour = 0;
        public readonly double _sickTotalHour = 0;

        public LeaveBalanceHelper() { }

        public LeaveBalanceHelper(DateTime workStartDate, DateTime ehcStartDate)
        {
            _workStartDate = workStartDate;
            _ehcStartDate = ehcStartDate;
            if (IsLeapYear(ehcStartDate.Year) && ehcStartDate.Month == 2 && ehcStartDate.Day == 29 && !IsLeapYear(_now.Year))
            {
                _currentEhcDate = new DateTime(_now.Year, ehcStartDate.Month, ehcStartDate.Day - 1);
                _currentWorkDate = new DateTime(_now.Year, workStartDate.Month, workStartDate.Day - 1);
            }
            else
            {
                _currentEhcDate = new DateTime(_now.Year, ehcStartDate.Month, ehcStartDate.Day);
                _currentWorkDate = new DateTime(_now.Year, workStartDate.Month, workStartDate.Day);
            }

            if (_ehcStartDate.Year == _now.Year)
            {
                var a1 = _currentEhcDate < _currentWorkDate ? CalculateTotalVacationHours(_currentWorkDate, _endDate) : CalculateTotalVacationHours(_currentEhcDate, _endDate);
                _annualTotalHour = DateTimeHelper.RoundToNearestHalf(a1);
            }
            else
            {
                var a1 = _currentEhcDate < _currentWorkDate ? CalculateTotalVacationHours(_currentWorkDate, _endDate) : CalculateTotalVacationHours(_currentEhcDate, _endDate);
                var a2 = _currentEhcDate < _currentWorkDate ? CalculateTotalVacationHours(_currentEhcDate, _currentWorkDate) : CalculateTotalVacationHours(_currentWorkDate, _currentEhcDate);
                var a3 = _currentEhcDate < _currentWorkDate ? CalculateTotalVacationHours(_startDate, _currentEhcDate) : CalculateTotalVacationHours(_startDate, _currentWorkDate);
                _annualTotalHour = DateTimeHelper.RoundToNearestHalf((a1 + a2 + a3)) == 119.5 ? 120.0 : DateTimeHelper.RoundToNearestHalf((a1 + a2 + a3));
            }
            var s1 = CalculateTotalSickHours(_currentEhcDate, _endDate);
            var s2 = CalculateTotalSickHours(_startDate, _currentEhcDate);
            _sickTotalHour = DateTimeHelper.RoundToNearestHalf(s1 + s2);
        }

        public double CalculateAnnualHours(DateTime workStartDate, DateTime ehcStartDate, DateTime? startDate, DateTime? endDate, out List<Atd_AnnualInfos> infos)
        {

            List<Atd_AnnualInfos> infoList = new();
            infos = infoList;
            startDate = startDate < ehcStartDate ? ehcStartDate : startDate;
            endDate = endDate < ehcStartDate ? ehcStartDate : endDate;
            if (startDate == endDate)
                return 0;

            List<DateTime> cDates = new List<DateTime>();
            cDates.Add(Convert.ToDateTime(startDate));
            for (DateTime date = Convert.ToDateTime(startDate); date < endDate; date = date.AddDays(1))
            {
                var year = date.Year;
                DateTime ehcDate, workDate;
                if (IsLeapYear(ehcStartDate.Year) && ehcStartDate.Month == 2 && ehcStartDate.Day == 29 && !IsLeapYear(year))
                {
                    ehcDate = new DateTime(year, ehcStartDate.Month, ehcStartDate.Day - 1);
                    workDate = new DateTime(year, workStartDate.Month, workStartDate.Day - 1);
                }
                else
                {
                    ehcDate = new DateTime(year, ehcStartDate.Month, ehcStartDate.Day);
                    workDate = new DateTime(year, workStartDate.Month, workStartDate.Day);
                }
                if (date == ehcDate || date == workDate)
                {
                    cDates.Add(date);
                }
            }
            cDates.Add(Convert.ToDateTime(endDate));
            double total = 0.0, hour;

            for (int i = 0; i < cDates.Count - 1; i++)
            {
                Atd_AnnualInfos info = new Atd_AnnualInfos();
                info.start = cDates[i].ToString("yyyy-MM-dd");
                info.end = cDates[i + 1].ToString("yyyy-MM-dd");
                info.day = (cDates[i + 1] - cDates[i]).Days;
                var ehcYear = CalculateTotalServiceYears(ehcStartDate, cDates[i]) + 1;
                var workYear = CalculateTotalServiceYears(workStartDate, cDates[i]);
                info.ehiYear = ehcYear;
                info.workYear = workYear;
                info.standard = GetYearHour(workYear, ehcYear)[ehcYear] / 8;
                hour = Convert.ToDouble(info.standard) * 8 * Convert.ToDouble(info.day) / 365;
                info.total = Math.Round(hour, 3);
                total += hour;
                if (info.day != 0)
                    infos.Add(info);
            }
            total = total == 119.5 ? 120 : total;
            return total;
        }

        public double CalculateTotalVacationHours(DateTime d1, DateTime d2)
        {
            var dayCount = (d2 - d1).Days;
            int _totalWorkYears = CalculateTotalServiceYears(_workStartDate, d1);
            int _totalServiceYears = CalculateTotalServiceYears(_ehcStartDate, d1) + 1;

            var serviceYearDic = GetYearHour(_totalWorkYears, _totalServiceYears);
            #region MyRegion
            //if (_totalWorkYears >= 1 && _totalWorkYears < 10)
            //{
            //    for (int i = 1; i <= _totalServiceYears; i++)
            //    {
            //        if (i <= 2)
            //        {
            //            serviceYearDic.Add(i, i * 5 * 8);
            //        }
            //        else if (i <= 7 && i >= 3)
            //        {
            //            serviceYearDic.Add(i, 80 + (i - 2) * 8);
            //        }
            //        else
            //        {
            //            serviceYearDic.Add(i, 120);
            //        }
            //    }
            //    //    serviceYearTable = new Dictionary<int, Tuple<double, int>>
            //    //{
            //    //    { 1, new Tuple<double, int>(5.0, 40) },
            //    //    { 2, new Tuple<double, int>(10.0, 80) },
            //    //    { 3, new Tuple<double, int>(11.0, 88) },
            //    //    { 4, new Tuple<double, int>(12.0, 96) },
            //    //    { 5, new Tuple<double, int>(13.0, 104) },
            //    //    { 6, new Tuple<double, int>(14.0, 112) },
            //    //    { 7, new Tuple<double, int>(15.0, 120) },
            //    //    { _totalServiceYears>7?_totalServiceYears:7,new Tuple<double, int>(15.0, 120) }
            //    //};

            //}
            //else if (_totalWorkYears >= 10 && _totalWorkYears < 20)
            //{
            //    for (int i = 1; i <= _totalServiceYears; i++)
            //    {
            //        if (i <= 2)
            //        {
            //            serviceYearDic.Add(i, 80);
            //        }
            //        else if (i <= 7 && i >= 3)
            //        {
            //            serviceYearDic.Add(i, 80 + (i - 2) * 8);
            //        }
            //        else
            //        {
            //            serviceYearDic.Add(i, 120);
            //        }
            //    }

            //}
            //else
            //{
            //    serviceYearDic.Add(_totalServiceYears, 120);
            //}
            #endregion

            int x = serviceYearDic[_totalServiceYears];
            var totalHour = Convert.ToDouble(serviceYearDic[_totalServiceYears]) * dayCount / 365;
            return totalHour;
            //return Math.Ceiling(totalHour);
        }

        public Dictionary<int, int> GetYearHour(int _totalWorkYears, int _totalServiceYears)
        {
            var serviceYearDic = new Dictionary<int, int>();
            if (_totalWorkYears < 1 && _totalWorkYears < 10)
            {
                serviceYearDic.Add(_totalServiceYears, 40);
            }
            else if (_totalWorkYears >= 1 && _totalWorkYears < 10)
            {
                for (int i = 1; i <= _totalServiceYears; i++)
                {
                    if (i <= 2)
                    {
                        serviceYearDic.Add(i, i * 5 * 8);
                    }
                    else if (i <= 7 && i >= 3)
                    {
                        serviceYearDic.Add(i, 80 + (i - 2) * 8);
                    }
                    else
                    {
                        serviceYearDic.Add(i, 120);
                    }
                }


            }
            else if (_totalWorkYears >= 10 && _totalWorkYears < 20)
            {
                for (int i = 1; i <= _totalServiceYears; i++)
                {
                    if (i <= 2)
                    {
                        serviceYearDic.Add(i, 80);
                    }
                    else if (i <= 7 && i >= 3)
                    {
                        serviceYearDic.Add(i, 80 + (i - 2) * 8);
                    }
                    else
                    {
                        serviceYearDic.Add(i, 120);
                    }
                }

            }
            else
            {
                serviceYearDic.Add(_totalServiceYears, 120);
            }
            return serviceYearDic;
        }

        public double CalculateTotalSickHours(DateTime d1, DateTime d2)
        {
            var dayCount = (d2 - d1).Days;
            int _totalServiceYears = CalculateTotalServiceYears(_ehcStartDate, d1) + 1;

            var serviceYearDic = new Dictionary<int, double>();
            var sickHour = _totalServiceYears < 2 ? 0 : 56;

            var totalHour = Convert.ToDouble(sickHour) * dayCount / 365;
            return totalHour;
            //return Math.Ceiling(totalHour);
        }

        public double CalculateTotalParentalHours(DateTime d1, DateTime d2, bool lastYear)
        {
            var dayCount = Convert.ToDouble((d2 - d1).Days);
            var totalHour = (8.0 * dayCount) / 365;
            return Math.Round(totalHour, MidpointRounding.AwayFromZero) * 8;//向上取整 0.1=1
            //if (!lastYear)
            //{
            //    return Math.Ceiling(totalHour) * 8;//向上取整 0.1=1
            //}
            //else
            //{
            //    return Math.Floor(totalHour) * 8;
            //}

        }

        private int CalculateTotalServiceYears(DateTime startDate, DateTime endDate)
        {
            var today = endDate.Date;
            var yearsPassed = today.Year - startDate.Year;
            if (startDate.Month > today.Month || startDate.Month == today.Month && startDate.Day > today.Day)
                --yearsPassed;
            //else
            //    ++yearsPassed;
            return yearsPassed;
        }


        public bool IsLeapYear(int year)
        {
            return year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
        }

    }
}
