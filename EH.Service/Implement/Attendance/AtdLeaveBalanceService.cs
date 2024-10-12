using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Interface.Attendance;
using EH.Service.Interface.Attendance;
using EH.Repository.Interface.Sys;
using System.Net.Sockets;
using System.Collections.Generic;
using static EH.System.Commons.LeaveBalanceHelper;
using EH.System.Models.Dtos;
using System.Data;
using Microsoft.EntityFrameworkCore;
using EH.Repository.DataAccess;
using Org.BouncyCastle.Crypto;

namespace EH.Service.Implement.Attendance
{
    public class AtdLeaveBalanceService : BaseService<Atd_LeaveBalance>, IAtdLeaveBalanceService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IAtdLeaveBalanceRepository repository;
        private readonly IAtdLeaveBalanceBakRepository bakRepository;
        private readonly IAtdLeaveFormRepository formRepository;
        private readonly ISysUsersRepository userRepository;
        private readonly IAtdLeaveSettingRepository settingRepository;
        public AtdLeaveBalanceService(IAtdLeaveFormRepository formRepository, IAtdLeaveSettingRepository settingRepository, IAtdLeaveBalanceRepository repository, ISysUsersRepository userRepository, IAtdLeaveBalanceBakRepository bakRepository, LogHelper logHelper) : base(repository, logHelper)
        {
            this.logHelper = logHelper;
            this.repository = repository;
            this.userRepository = userRepository;
            this.settingRepository = settingRepository;
            this.formRepository = formRepository;
            this.bakRepository = bakRepository;
        }

        public List<Atd_BalanceStatics> Statistics(PageRequest request, out int totalCount)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("UserName");
            dt.Columns.Add("ChineseName");
            dt.Columns.Add("Department");
            dt.Columns.Add("AnnualTotal");
            dt.Columns.Add("AnnualAvailable");
            dt.Columns.Add("AnnualCarryover");
            dt.Columns.Add("AnnualUsed");
            dt.Columns.Add("AnnualLocked");
            dt.Columns.Add("AnnualCarryoverUsed");

            dt.Columns.Add("Sick");
            dt.Columns.Add("Personal");
            dt.Columns.Add("Paternity");
            dt.Columns.Add("Parental");
            dt.Columns.Add("Nursing");

            dt.Columns.Add("Abortion");
            dt.Columns.Add("Bereavement");
            dt.Columns.Add("Maternity");
            dt.Columns.Add("Wedding");
            dt.Columns.Add("Marriage");
            dt.Columns.Add("Prenatal");
            dt.Columns.Add("PrenatalCheckup");

            // 预处理
            var requestWhere = request.where.Split('&');
            var requestDefaultWhere = request.defaultWhere.Split(',');

            var userName = requestWhere[0].ToLower();
            var department = requestWhere[1].ToUpper();
            var userId = requestDefaultWhere[0].Split('=')[1];
            var isHR = Convert.ToBoolean(requestDefaultWhere[1].Split('=')[1]);
            var isNormal = Convert.ToBoolean(requestWhere[2]);

            var currentYear = DateTime.Now.Year;

            // 获取用户数据和设置
            var allUsers = userRepository.Entities.ToList();
            var settingList = settingRepository.Entities.ToList();

            // 快速查找用户
            var user = allUsers.FirstOrDefault(i => i.UserName == userId);
            var userList = allUsers.Where(i =>
                (user.UserName == i.Report || user.FullName.Contains(i.Report)) ||
                (user.UserName == i.LastReport || user.FullName.Contains(i.LastReport))
            ).Distinct().ToList();

            // 处理 CS 管理员
            if (new[] { "nancyl", "marsw", "Nancy Lin", "Mars Wan" }.Contains(user.UserName))
            {
                var csManager = new[] { "nancyl", "marsw", "Nancy Lin", "Mars Wan" };
                userList = allUsers.Where(i =>
                    csManager.Contains(i.Report) || csManager.Contains(i.LastReport)
                ).Distinct().ToList();
            }

            // 插入当前用户
            if (!userList.Any(i => i.UserName == user.UserName))
            {
                userList.Insert(0, user);
            }
            else
            {
                userList = userList.Where(i => i.UserName != user.UserName).ToList();
                userList.Insert(0, user);
            }

            // 根据条件过滤
            if (isHR)
            {
                userList = allUsers;
            }
            else if (isNormal)
            {
                userList = userList.Where(i => i.UserName == user.UserName).ToList();
            }

            if (!string.IsNullOrEmpty(userName))
            {
                userList = userList.Where(i => i.UserName.ToLower().Contains(userName) || i.FullName.ToLower().Contains(userName)).ToList();
            }
            if (!string.IsNullOrEmpty(department))
            {
                userList = userList.Where(i => i.Department.ToUpper().Contains(department)).ToList();
            }
            var balanceList = repository.Entities.ToList();
            var formList = formRepository.Entities.ToList();

            // 处理数据
            Parallel.ForEach(userList, u =>
            //userList.ForEach(u =>
                 {
                     var row = dt.NewRow();
                     row["UserName"] = u.FullName;
                     row["ChineseName"] = u.ChineseName;
                     row["Department"] = u.Department;

                     foreach (var s in settingList)
                     {
                         if (s.LeaveType == "annual")
                         {
                             var annual = balanceList.FirstOrDefault(i => i.LeaveType == "annual" && i.UserId == u.UserName && i.Year == currentYear);
                             if (annual != null)
                             {
                                 row["AnnualTotal"] = annual.Total;
                                 row["AnnualAvailable"] = annual.Available;
                                 row["AnnualUsed"] = annual.Used;
                                 row["AnnualLocked"] = annual.Locked;
                                 row["AnnualCarryOver"] = annual.AVCarryoverTotal;

                                 // 打印出所有相关值
                                 //Console.WriteLine($"annual.Overdue: {annual.Overdue}, annual.Used: {annual.Used}, annual.Locked: {annual.Locked}");

                                 // 计算并打印 AnnualCarryoverUsed
                                 var carryoverUsed = Math.Max(0, (annual.AVCarryoverTotal) - (annual.Used) - (annual.Locked));
                                 row["AnnualCarryoverUsed"] = carryoverUsed;
                             }
                             else
                             {
                                 row["AnnualTotal"] = row["AnnualAvailable"] = row["AnnualUsed"] = row["AnnualLocked"] = row["AnnualCarryOver"] = row["AnnualCarryoverUsed"] = 0;
                             }
                         }
                         else
                         {
                             var leaveSum = formList.Where(i => i.UserId == u.UserName && i.LeaveType == s.LeaveType && i.CurrentState == 2 && i.StartDate.Year == currentYear).Sum(i => i.TotalHours);
                             row[s.LeaveType] = leaveSum;
                         }
                     }

                     lock (dt)
                     {
                         dt.Rows.Add(row);
                     }
                 });

            // 分页处理
            var result = dt.ToObject<List<Atd_BalanceStatics>>();
            totalCount = result.Count;
            var res = result.Skip((request.PageIndex - 1) * request.PageSize).Take(request.PageSize).ToList();
            return res;

            //var requestWhere = request.where.Split('&');
            //var requestDefaultWhere = request.defaultWhere.Split(',');

            //var userName = requestWhere[0];
            //var department = requestWhere[1];
            //var userId = requestDefaultWhere[0].Split('=')[1];
            //var isHR = Convert.ToBoolean(requestDefaultWhere[1].Split('=')[1]);
            //var isNormal = Convert.ToBoolean(requestWhere[2]);

            //var currentYear = DateTime.Now.Year;
            //var allUsers = userRepository.Entities.ToList();
            //var user = allUsers.FirstOrDefault(i => i.UserName == userId);
            //var userList = allUsers.Where(i =>
            //    (user.UserName == i.Report || user.FullName.Contains(i.Report)) ||
            //    (user.UserName == i.LastReport || user.FullName.Contains(i.LastReport))
            //).Distinct().ToList();

            //if (user.UserName == "nancyl" || user.UserName == "marsw")
            //{
            //    var csManager = new[] { "nancyl", "marsw", "Nancy Lin", "Mars Wan" };
            //    userList = allUsers.Where(i =>
            //        csManager.Contains(i.Report) || csManager.Contains(i.LastReport)
            //    ).Distinct().ToList();
            //}

            //if (!userList.Any(i => i.UserName == user.UserName))
            //{
            //    userList.Insert(0, user);
            //}
            //else
            //{
            //    userList = userList.Where(i => i.UserName != user.UserName).ToList();
            //    userList.Insert(0, user);
            //}

            //if (isHR)
            //{
            //    userList = allUsers;
            //}
            //else if (isNormal)
            //{
            //    userList = userList.Where(i => i.UserName == user.UserName).ToList();
            //}

            //if (!string.IsNullOrEmpty(userName))
            //{
            //    userName = userName.ToLower();
            //    userList = userList.Where(i => i.UserName.ToLower().Contains(userName) || i.FullName.ToLower().Contains(userName)).ToList();
            //}
            //if (!string.IsNullOrEmpty(department))
            //{
            //    department = department.ToUpper();
            //    userList = userList.Where(i => i.Department.ToUpper().Contains(department)).ToList();
            //}

            //var settingList = settingRepository.Entities.ToList();

            //foreach (var u in userList)
            //{
            //    dt.Rows.Add();
            //    var row = dt.Rows[dt.Rows.Count - 1];
            //    row["UserName"] = u.FullName;
            //    row["ChineseName"] = u.ChineseName;
            //    row["Department"] = u.Department;

            //    foreach (var s in settingList)
            //    {
            //        if (s.LeaveType == "annual")
            //        {
            //            var annual = repository.FirstOrDefault(i => i.LeaveType == "annual" && i.UserId == u.UserName && i.Year == currentYear);
            //            if (annual != null)
            //            {
            //                row["AnnualTotal"] = annual.Total;
            //                row["AnnualAvailable"] = annual.Available;
            //                row["AnnualUsed"] = annual.Used;
            //                row["AnnualLocked"] = annual.Locked;
            //                row["AnnualCarryOver"] = annual.Overdue;
            //                row["AnnualCarryoverUsed"] = Math.Max(0, (decimal)(annual.Overdue - annual.Used));
            //            }
            //            else
            //            {
            //                row["AnnualTotal"] = row["AnnualAvailable"] = row["AnnualUsed"] = row["AnnualLocked"] = row["AnnualCarryOver"] = row["AnnualCarryoverUsed"] = 0;
            //            }
            //        }
            //        else
            //        {
            //            var leaveSum = formRepository.Where(i => i.UserId == u.UserName && i.LeaveType == s.LeaveType && i.CurrentState == 2 && i.StartDate.Year == currentYear).Sum(i => i.TotalHours);
            //            row[s.LeaveType] = leaveSum;
            //        }
            //    }
            //}

            //var result = dt.ToObject<List<Atd_BalanceStatics>>();
            //totalCount = result.Count;
            //var res = result.Skip((request.PageIndex - 1) * request.PageSize).Take(request.PageSize).ToList();
            //return res;

            //var userName = request.where.Split('&')[0];
            //var department = request.where.Split('&')[1];

            //var userId = request.defaultWhere.Split(',')[0].Split('=')[1];
            //var isHR = Convert.ToBoolean(request.defaultWhere.Split(',')[1].Split('=')[1]);
            //var isNormal = Convert.ToBoolean(request.where.Split('&')[2]);

            //var currentYear = DateTime.Now.Year;
            //var user = userRepository.Entities.FirstOrDefault(i => i.UserName == userId);
            //var userList = userRepository.Entities.Where(i => (user.UserName == i.Report || user.FullName.Contains(i.Report)) || (user.UserName == i.LastReport || user.FullName.Contains(i.LastReport))).Distinct().ToList();
            //if (user.UserName == "nancyl" || user.UserName == "marsw")
            //{
            //    var csManager = new string[] { "nancyl", "marsw", "Nancy Lin", "Mars Wan" };
            //    userList = userRepository.Entities.Where(i => (csManager.Contains(i.Report) || (csManager.Contains(i.LastReport)))).Distinct().ToList();
            //}

            //if (userList.Where(i => i.UserName == user.UserName).Count() <= 0)
            //    userList.Insert(0, user);
            //else
            //{
            //    userList = userList.Where(i => i.UserName != user.UserName).ToList();
            //    userList.Insert(0, user);
            //}
            //if (isHR)
            //{
            //    userList = userRepository.Entities.ToList();
            //}
            //else if (isNormal)
            //{
            //    userList = userList.Where(i => i.UserName == user.UserName).ToList();
            //}
            //else
            //{

            //}

            //if (!string.IsNullOrEmpty(userName))
            //{
            //    userList = userList.Where(i => i.UserName.Contains(userName) || i.FullName.Contains(userName)).ToList();
            //}
            //if (!string.IsNullOrEmpty(department))
            //{
            //    department = department.ToUpper();
            //    userList = userList.Where(i => i.Department.ToUpper().Contains(department)).ToList();
            //}



            //foreach (var u in userList)
            //{
            //    //var settingList=  settingRepository.Where(i => i.Qualification == "0" || i.Qualification == u.Gender.ToString());
            //    var settingList = settingRepository.Entities.ToList();
            //    dt.Rows.Add();
            //    dt.Rows[dt.Rows.Count - 1]["UserName"] = u.FullName;
            //    dt.Rows[dt.Rows.Count - 1]["ChineseName"] = u.ChineseName;
            //    dt.Rows[dt.Rows.Count - 1]["Department"] = u.Department;
            //    foreach (var s in settingList)
            //    {
            //        if (s.LeaveType == "annual")
            //        {
            //            var annual = repository.FirstOrDefault(i => i.LeaveType == "annual" && i.UserId == u.UserName && i.Year == currentYear);
            //            dt.Rows[dt.Rows.Count - 1]["AnnualTotal"] = annual != null ? annual.Total : 0;
            //            dt.Rows[dt.Rows.Count - 1]["AnnualAvailable"] = annual != null ? annual.Available : 0;
            //            dt.Rows[dt.Rows.Count - 1]["AnnualUsed"] = annual != null ? annual.Used : 0;
            //            dt.Rows[dt.Rows.Count - 1]["AnnualLocked"] = annual != null ? annual.Locked : 0;
            //            dt.Rows[dt.Rows.Count - 1]["AnnualCarryOver"] = annual != null ? annual.Overdue : 0;
            //            var carryoverUsed = annual != null ? annual.Overdue - annual.Used >= 0 ? annual.Overdue - annual.Used : 0 : 0;
            //            dt.Rows[dt.Rows.Count - 1]["AnnualCarryoverUsed"] = carryoverUsed;
            //        }
            //        else
            //        {
            //            var leaveSum = formRepository.Where(i => i.UserId == u.UserName && i.LeaveType == s.LeaveType && i.CurrentState == 2 && i.StartDate.Year == currentYear).Sum(i => i.TotalHours);
            //            dt.Rows[dt.Rows.Count - 1][s.LeaveType] = leaveSum;
            //        }
            //    }
            //}

            //var result = dt.ToObject<List<Atd_BalanceStatics>>();
            //totalCount = result.Count;
            //var res = result.Skip((request.PageIndex - 1) * (request.PageSize)).Take(request.PageSize).ToList();
            //return res;
        }
        public List<Atd_AnnualInfos> GetInfo(Atd_AnnualInfoReq annualInfo)
        {
            var user = userRepository.Entities.FirstOrDefault(i => i.UserName == annualInfo.userId);
            if (user == null)
            {
                logHelper.LogError("GetInfo User NotFound");
                return null;
            }
            user.StartWorkDate = user.StartWorkDate == null ? user.EhiStratWorkDate : user.StartWorkDate;
            var helper = new LeaveBalanceHelper((DateTime)user.StartWorkDate, (DateTime)user.EhiStratWorkDate);
            annualInfo.StartDate = annualInfo.StartDate ?? new DateTime(DateTime.Now.Year, 1, 1);
            annualInfo.EndDate = annualInfo.EndDate ?? new DateTime(DateTime.Now.Year, 12, 31);
            var res = helper.CalculateAnnualHours((DateTime)user.StartWorkDate, (DateTime)user.EhiStratWorkDate, annualInfo.StartDate, annualInfo.EndDate, out List<Atd_AnnualInfos> infos);
            var result = infos;
            return result;
        }

        public bool CalculateAnnualAndSick(List<string> userIds)
        {
            List<Atd_LeaveBalance> annualBalances = new();
            List<Atd_LeaveBalance> sickBalances = new();
            try
            {
                foreach (var userId in userIds)
                {
                    var user = userRepository.Entities.FirstOrDefault(i => i.UserName == userId);
                    if (user == null)
                    {
                        logHelper.LogError("CalculateAnnualAndSick User NotFound");
                        continue;
                    }
                    user.StartWorkDate = user.StartWorkDate == null ? user.EhiStratWorkDate : user.StartWorkDate;
                    var helper = new LeaveBalanceHelper((DateTime)user.StartWorkDate, (DateTime)user.EhiStratWorkDate);
                    var year = DateTime.Now.Year;

                    var sick = helper._sickTotalHour;
                    var annual = helper._annualTotalHour;

                    var annualBalance = repository.FirstOrDefault(i => i.UserId == user.UserName && i.LeaveType == "annual" && i.Year == year);
                    var preAnnualBalance = repository.FirstOrDefault(i => i.UserId == user.UserName && i.LeaveType == "annual" && i.Year == DateTime.Now.AddYears(-1).Year);
                    var sickBalance = repository.FirstOrDefault(i => i.UserId == user.UserName && i.LeaveType == "sick" && i.Year == year);

                    repository.BeginTransaction();

                    if (annualBalance == null)
                    {
                        Atd_LeaveBalance atd_LeaveBalance = new()
                        {
                            UserId = userId,
                            LeaveType = "annual",
                            Year = year,
                            Total = Convert.ToDecimal(annual) + (preAnnualBalance == null ? 0 : preAnnualBalance.Available),
                            Available = Convert.ToDecimal(annual) + (preAnnualBalance == null ? 0 : preAnnualBalance.Available),
                            Used = Convert.ToDecimal(0),
                            Locked = Convert.ToDecimal(0),
                            AVCarryoverTotal = (preAnnualBalance == null ? 0 : preAnnualBalance.Available),
                            Remark = $"{year} annual",
                        };
                        annualBalances.Add(atd_LeaveBalance);
                    }
                    if (sickBalance == null)
                    {
                        Atd_LeaveBalance atd_LeaveBalance = new()
                        {
                            UserId = userId,
                            LeaveType = "sick",
                            Year = year,
                            Total = Convert.ToDecimal(sick),
                            Available = Convert.ToDecimal(sick),
                            Used = Convert.ToDecimal(0),
                            Locked = Convert.ToDecimal(0), 
                            Remark = $"{year} sick",
                        };
                        sickBalances.Add(atd_LeaveBalance);
                    }

                }

                repository.BeginTransaction();
                repository.AddRange(annualBalances);
                repository.AddRange(sickBalances);
                repository.Commit();
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError("CalculateAnnualAndSick error" + ex.ToString());
                return false;
            }
        }

        public bool CalculatePersonal(List<string> userIds)
        {
            try
            {
                List<Atd_LeaveBalance> atd_LeaveBalances = new List<Atd_LeaveBalance>();
                foreach (var userId in userIds)
                {
                    var user = userRepository.Entities.FirstOrDefault(i => i.UserName == userId);
                    if (user == null)
                    {
                        logHelper.LogError("CalculatePensonal User NotFound");
                        return false;
                    }
                    var year = DateTime.Now.Year;
                    var personalBalance = repository.FirstOrDefault(i => i.UserId == user.UserName && i.LeaveType == "personal" && i.Year == year);

                    if (personalBalance == null)
                    {
                        Atd_LeaveBalance atd_LeaveBalance = new Atd_LeaveBalance()
                        {
                            UserId = userId,
                            LeaveType = "personal",
                            Year = year,
                            Total = Convert.ToDecimal(120),
                            Available = Convert.ToDecimal(120),
                            Used = Convert.ToDecimal(0),
                            Locked = Convert.ToDecimal(0), 
                            Remark = $"{year} personal",
                        };
                        atd_LeaveBalances.Add(atd_LeaveBalance);
                    }
                }
                repository.BeginTransaction();
                repository.AddRange(atd_LeaveBalances);
                repository.Commit();
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError("CalculatePensonal error" + ex.ToString());
                return false;
            }

        }

        public bool CalculateParentalAndBreastfeeding(string userId, DateTime bornDate)
        {
            try
            {
                var user = userRepository.Entities.FirstOrDefault(i => i.UserName == userId);
                if (user == null)
                {
                    logHelper.LogError("CalculateParentalAndBreastfeeding User NotFound");
                    return false;
                }

                var parentalBalances = repository.Where(i => i.UserId == userId && i.LeaveType == "parental").ToList();

                var beginYear = bornDate.Year;
                var fullDate = bornDate.AddYears(3);
                var endYear = fullDate.Year;

                var helper = new LeaveBalanceHelper();
                var firstYearDay = helper.CalculateTotalParentalHours(bornDate, Convert.ToDateTime(beginYear.ToString() + "-12-31"), false);
                var lastYearDay = helper.CalculateTotalParentalHours(Convert.ToDateTime(endYear.ToString() + "-01-01"), fullDate, true);

                List<Atd_LeaveBalance> addLbs = new();
                List<Atd_LeaveBalance> updateLbs = new();
                for (int i = 0; i < 4; i++)
                {
                    Atd_LeaveBalance atd_LeaveBalance = new Atd_LeaveBalance()
                    {
                        UserId = userId,
                        LeaveType = "parental",
                        Year = beginYear + i,
                        Total = Convert.ToDecimal(i == 0 ? firstYearDay : i == 3 ? lastYearDay : 8 * 8),
                        Available = Convert.ToDecimal(i == 0 ? firstYearDay : i == 3 ? lastYearDay : 8 * 8),
                        Used = Convert.ToDecimal(0),
                        Locked = Convert.ToDecimal(0), 
                        Remark = $"{beginYear + i} parental",
                    };
                    var yearLb = parentalBalances.FirstOrDefault(p => p.Year == (beginYear + i));
                    if (yearLb == null)
                    {
                        addLbs.Add(atd_LeaveBalance);
                    }
                    else
                    {
                        yearLb.Total = atd_LeaveBalance.Total;
                        yearLb.Available = atd_LeaveBalance.Available;
                        updateLbs.Add(yearLb);
                    }
                }

                repository.BeginTransaction();
                repository.AddRange(addLbs);
                repository.UpdateRange(updateLbs);
                repository.Commit();
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError("CalculateParentalAndBreastfeeding error" + ex.ToString());
                return false;
            }
        }

        public string ClearCarryoverAnnual(bool isClear=false)
        {
            try
            {
                var needClearBalance = repository.Where(i => i.LeaveType == "annual" && i.Year == DateTime.Now.Year&&i.IsClear==false).ToList();
                var needClearList = needClearBalance.ToObject<List<Atd_LeaveBalance_Bak>>();
                var ids = bakRepository.Entities.Select(i => i.ID);
                //把旧的写入备份表.Except(needClearList.Where(i => i.AnnualClear != 0))
                bakRepository.UpdateRange(needClearList.Where(i => ids.Contains(i.ID)));//存在更新
                bakRepository.AddRange(needClearList.Where(i => !ids.Contains(i.ID)));//不存在添加

                List<Atd_LeaveBalance> baks = new();
                needClearBalance.ForEach(e =>
                {
                    var carryoverHour = e.AVCarryoverTotal;
                    var used = e.Used;
                    var locked = e.Locked;
                    var carryoverUnused = Math.Max(0, carryoverHour - used);
                    //if (carryoverUnused == 0)
                    //{
                    //    e.AnnualClear = 0;
                    //}
                    //else if (locked < carryoverUnused)
                    //{
                    //    baks.Add(e);//添加到导出excel的数据里
                    //}
                    //else
                    //{
                    //    e.AnnualClear = 0;//之后做优化
                       
                    //}

                    var usedHour = e.Used + e.Locked;


                    if (usedHour >= carryoverHour)//如果大于则继续
                    {
                        e.AVCarryoverCleared = 0;
                    }
                    else
                    {
                        baks.Add(e);//添加到导出excel的数据里
                        var clearHour = carryoverHour - usedHour;
                        e.AVCarryoverCleared = clearHour;
                        e.Available -= clearHour;
                    }
                });

                if (baks.Count <= 0||isClear)//如果导出数据不为空，则返回excel数据
                {
                    needClearBalance.ForEach(i => { i.IsClear = true; });
                    repository.UpdateRange(needClearBalance);
                    return "true";
                }
                else//清理
                {
                    var bytes = OfficeHelper.ExportToExcel(baks);
                    return Convert.ToBase64String(bytes);
                }
            }
            catch (Exception ex)
            {
                logHelper.LogError("clearAnnualfail:" + ex.ToString());
                return "false";
            }



        }

    }
}
