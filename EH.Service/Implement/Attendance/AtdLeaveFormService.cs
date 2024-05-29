using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Interface.Attendance;
using EH.Service.Interface.Attendance;
using EH.Repository.Interface.Sys;
using System.Collections.Generic;
using EH.System.Models.Dtos;
using EH.System.Models.Common;
using NPOI.POIFS.FileSystem;
using System.Reflection.Metadata.Ecma335;
using Microsoft.EntityFrameworkCore.Storage;
using EH.Repository.Implement.Sys;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using static System.Net.WebRequestMethods;
using EH.Repository.Implement;
using NPOI.SS.Formula.Functions;
using System.Numerics;
using NPOI.Util;
using NPOI.XWPF.UserModel;

namespace EH.Service.Implement.Attendance
{


    public class AtdLeaveFormService : BaseService<Atd_LeaveForm>, IAtdLeaveFormService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IAtdLeaveFormRepository repository;
        private readonly IAtdLeaveSettingRepository settingRepository;
        private readonly IAtdLeaveProcessService processService;
        private readonly IAtdLeaveProcessRepository processRepository;
        private readonly IAtdLeaveBalanceRepository balanceRepository;
        private readonly ISysUsersRepository userRepository;
        private readonly IAtdHolidaySettingService holidayService;

        private readonly IMailService mailService;
        private readonly IConfiguration configuration;

        public AtdLeaveFormService(IAtdLeaveFormRepository repository, IAtdLeaveProcessService processService, IAtdLeaveSettingRepository settingRepository, IAtdLeaveProcessRepository processRepository, ISysUsersRepository userRepository, IMailService mailService, IAtdLeaveBalanceRepository balanceRepository, IAtdHolidaySettingService holidayService, IConfiguration configuration, LogHelper logHelper) : base(repository, logHelper)
        {
            this.logHelper = logHelper;
            this.repository = repository;
            this.processService = processService;
            this.settingRepository = settingRepository;
            this.processRepository = processRepository;
            this.userRepository = userRepository;
            this.mailService = mailService;
            this.balanceRepository = balanceRepository;
            this.holidayService = holidayService;
            this.configuration = configuration;
        }

        public List<Atd_FormWithState> QueryGetPageList(PageRequest<Atd_LeaveForm> request, out int totalCount)
        {
            List<Atd_FormWithState> result = new List<Atd_FormWithState>();
            var defaultList = new List<Atd_LeaveForm>();
            var userId = request.defaultWhere.Split(',')[0].Split('=')[1];
            var isHR = Convert.ToBoolean(request.defaultWhere.Split(',')[1].Split('=')[1]);
            var user = userRepository.Entities.FirstOrDefault(i => i.UserName == userId);
            var userList = userRepository.Entities.Where(i => (user.UserName == i.Report || user.FullName.Contains(i.Report)) || (user.UserName == i.LastReport || user.FullName.Contains(i.LastReport))).ToList().Distinct();

            if (isHR)
            {
                userList = userRepository.Entities.ToList();
            }

            request.defaultWhere = null;

            var whereCondition = request.GetWhere().Compile();
            var orderCondition = request.GetOrder().Compile();

            foreach (var u in userList)
            {
                var forms = repository.Entities.Where(i => i.UserId == u.UserName && i.CurrentState != -1).ToList();
                defaultList.AddRange(forms);
            }

            defaultList = defaultList.Where(whereCondition).ToList();
            totalCount = defaultList.Count;
            if (!request.isDesc)
            {
                defaultList = defaultList.OrderBy(orderCondition).Skip((request.PageIndex - 1) * request.PageSize).Take(request.PageSize).ToList();
            }
            else
            {
                defaultList = defaultList.OrderByDescending(orderCondition).Skip((request.PageIndex - 1) * request.PageSize).Take(request.PageSize).ToList();
            }
            defaultList.ForEach(i =>
            {
                Atd_FormWithState entity = i.ToObject<Atd_FormWithState>();
                var formUser = userRepository.FirstOrDefault(i => i.UserName == entity.UserId);
                entity.FullName = formUser.FullName;
                entity.ChineseName = formUser.ChineseName;
                var leaveProcesses = processRepository.FirstOrDefault(p => p.LeaveId == i.LeaveId && p.IsLastNode);
                if (leaveProcesses.ProcessState == "wait")
                {
                    entity.CurrentStep = "To";
                    var currentProcess = processRepository.Where(p => p.LeaveId == i.LeaveId && p.ProcessState == "wait").OrderBy(i => i.OrderNo).FirstOrDefault();
                    entity.CurrentOwner = userRepository.FirstOrDefault(i => i.UserName == currentProcess.UserId || i.FullName.Contains(currentProcess.UserId)).FullName;
                }
                else
                {
                    entity.CurrentStep = "End";
                    entity.CurrentOwner = userRepository.FirstOrDefault(i => i.UserName == leaveProcesses.UserId || i.FullName.Contains(leaveProcesses.UserId)).FullName;
                }
                result.Add(entity);
            });
            return result;
        }

        public List<Atd_FormWithState> GetPageListWithState(PageRequest<Atd_LeaveForm> request, out int totalCount)
        {
            List<Atd_FormWithState> result = new List<Atd_FormWithState>();
            var res = base.GetPageList(request, out totalCount);
            res.ForEach(i =>
            {
                Atd_FormWithState entity = i.ToObject<Atd_FormWithState>();
                if (i.CurrentState != -1)
                {
                    var leaveProcesses = processRepository.FirstOrDefault(p => p.LeaveId == i.LeaveId && p.IsLastNode);
                    if (leaveProcesses.ProcessState == "wait")
                    {
                        entity.CurrentStep = "To";
                        var currentProcess = processRepository.Where(p => p.LeaveId == i.LeaveId && p.ProcessState == "wait").OrderBy(i => i.OrderNo).FirstOrDefault();
                        entity.CurrentOwner = userRepository.FirstOrDefault(i => i.UserName == currentProcess.UserId || i.FullName.Contains(currentProcess.UserId)).FullName;
                    }
                    else
                    {
                        entity.CurrentStep = "End";
                        entity.CurrentOwner = userRepository.FirstOrDefault(i => i.UserName == leaveProcesses.UserId || i.FullName.Contains(leaveProcesses.UserId)).FullName;
                    }

                }
                result.Add(entity);
            });
            return result;
        }

        public List<Atd_Statics> GetStatistic(PageRequest request, out int totalCout)
        {
            var users = userRepository.Entities;
            var forms = repository.Entities;

            var userName = request.where.Split('&')[0];
            var department = request.where.Split('&')[1];
            var leaveType = request.where.Split('&')[2];
            var startDate = Convert.ToDateTime(request.defaultWhere.Split('&')[0]);
            var endDate = Convert.ToDateTime(request.defaultWhere.Split('&')[1]);

            var result = (from u in users
                          join f in forms on
                          new { UserId = u.UserName, IsDeleted = false, CurrentState = 2 } equals
                          new { f.UserId, f.IsDeleted, f.CurrentState }
                          into temp
                          from t in temp.DefaultIfEmpty()
                          where t.StartDate >= startDate && t.EndDate <= endDate
                          && u.UserName.Contains(userName) && u.Department.Contains(department)
                          group new { u, t } by new { u.UserName, u.Department } into g

                          select new Atd_Statics
                          {
                              UserName = g.Key.UserName,
                              Department = g.Key.Department,
                              Annual = g.Sum(x => x.t.LeaveType == "annual" ? x.t.TotalHours : 0),
                              Sick = g.Sum(x => x.t.LeaveType == "sick" ? x.t.TotalHours : 0),
                              Personal = g.Sum(x => x.t.LeaveType == "personal" ? x.t.TotalHours : 0),
                              Wedding = g.Sum(x => x.t.LeaveType == "wedding" ? x.t.TotalHours : 0),
                              Paternity = g.Sum(x => x.t.LeaveType == "paternity" ? x.t.TotalHours : 0),
                              Parental = g.Sum(x => x.t.LeaveType == "parental" ? x.t.TotalHours : 0),
                              Breastfeeding = g.Sum(x => x.t.LeaveType == "breastfeeding" ? x.t.TotalHours : 0),
                              Prenatal = g.Sum(x => x.t.LeaveType == "prenatal" ? x.t.TotalHours : 0),
                              Abortion = g.Sum(x => x.t.LeaveType == "abortion" ? x.t.TotalHours : 0),
                              Bereavement = g.Sum(x => x.t.LeaveType == "bereavement" ? x.t.TotalHours : 0),
                              Maternity = g.Sum(x => x.t.LeaveType == "maternity" ? x.t.TotalHours : 0),
                          });

            if (!string.IsNullOrEmpty(leaveType))
            {
                result.Where(i => i.GetType().GetProperties().Select(p => p.Name.ToLower()).Contains(leaveType));
            }

            totalCout = result.Count();

            var res = result.Skip((request.PageIndex - 1) * request.PageSize).Take(request.PageSize).ToList();

            return res;
        }

        public List<decimal> GetHomePageBodyData(string userId)
        {
            Dictionary<string, object> obj = new Dictionary<string, object>();
            var forms = repository.Entities;
            var setting = settingRepository.Entities;

            var year = DateTime.Now.Year;
            var months = Enumerable.Range(1, 12);

            var filteredForms = forms.AsEnumerable()
                .Where(l => l.StartDate.Year == year && (l.CurrentState == 2 && !l.IsCancel) && l.UserId == userId)
                .ToList();

            var result = months.Select(m => new
            {
                month = m,
                hour = filteredForms.Where(l => l.StartDate.Month == m).Select(l => l.TotalHours).DefaultIfEmpty(0).Sum()
            })
            .OrderBy(x => x.month)
            .Select(x => x.hour ?? 0)
            .ToList();
            return result;



        }

        public Dictionary<string, int> GetHomePageData(string userId)
        {

            Dictionary<string, int> obj = new Dictionary<string, int>();
            var forms = repository.Entities;
            var process = processRepository.Entities;
            PageRequest<Atd_LeaveForm> pageRequest = new PageRequest<Atd_LeaveForm>()
            {
                PageIndex = 1,
                PageSize = 10,
                defaultWhere = $"userId={userId}"
            };
            var needApprovalForm = GetWaitAuditForm(pageRequest, out int totalCount);
            var needCount = totalCount;
            obj.Add("我的任务", needCount);
            var year = DateTime.Now.Year;
            var formDrafts = repository.Where(i => i.UserId == userId && i.CurrentState == -1).Where(i => i.CreateDate.Year == year).Count();
            obj.Add("草稿假单", formDrafts);
            var formProcess = repository.Where(i => i.UserId == userId && i.CurrentState == 0).Where(i => i.CreateDate.Year == year).Count();
            obj.Add("我的假单", formProcess);
            var formWait = repository.Where(i => i.UserId == userId && i.CurrentState == 1).Where(i => i.CreateDate.Year == year).Count();
            obj.Add("我的待批假单", formWait);
            var formDeny = repository.Where(i => i.UserId == userId && i.CurrentState == 3).Where(i => i.CreateDate.Year == year).Count();
            obj.Add("未通过假单", formDeny);

            return obj;
        }

        public bool Treate(List<string> ids)
        {
            try
            {
                foreach (var id in ids)
                {
                    var entity = repository.FirstOrDefault(i => i.ID.ToString() == id);
                    entity.IsTreated = !entity.IsTreated;
                    repository.Update(entity);
                }
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError("Treate error:" + ex.ToString());
                return false;
            }
        }

        public Atd_LeaveForm Save(Atd_LeaveForm entity)
        {
            var settingEntity = settingRepository.Entities.FirstOrDefault(i => i.LeaveType == entity.LeaveType);
            if (settingEntity == null) return null;

            var workTime = userRepository.FirstOrDefault(i => i.UserName == entity.UserId).Worktime;
            workTime = string.IsNullOrEmpty(workTime) ? "8:30-17:30" : workTime;
            var isContainHoliday = settingEntity.IsContainHoliday;
            var isNursing = settingEntity.LeaveType == "nursing";
            var totalHour = Convert.ToDecimal(CalculateLeaveHours(entity.StartDate, entity.EndDate, workTime, isContainHoliday, isNursing));

            entity.TotalHours = totalHour;
            // Convert.ToDecimal(CalculateTotalHours(entity.StartDate, entity.EndDate));

            return base.Insert(entity);
        }

        public JsonResultModel<Atd_LeaveForm> Apply(Atd_LeaveForm entity)
        {
            var settingEntity = settingRepository.Entities.FirstOrDefault(i => i.LeaveType == entity.LeaveType);
            if (settingEntity == null)
                return new JsonResultModel<Atd_LeaveForm>()
                {
                    Code = "100",
                    Result = null,
                    Message = "fail"
                };
            var workTime = userRepository.FirstOrDefault(i => i.UserName == entity.UserId).Worktime;
            workTime = string.IsNullOrEmpty(workTime) ? "8:30-17:30" : workTime;
            var isContainHoliday = settingEntity.IsContainHoliday;
            var isNursing = settingEntity.LeaveType == "nursing";
            var totalHour = Convert.ToDecimal(CalculateLeaveHours(entity.StartDate, entity.EndDate, workTime, isContainHoliday, isNursing));

            entity.TotalHours = totalHour;
            //Convert.ToDecimal(CalculateTotalHours(entity.StartDate, entity.EndDate));

            var result = GoProcess(entity);
            return result;
        }

        public override bool Delete(Atd_LeaveForm entity)
        {
            if (entity.IsCancel)//如果是销假
            {
                repository.Update(entity);
                var res = processService.AddProcessByLeaveType(entity.LeaveId);
                if (res)
                {
                    entity.CurrentState = 1;
                    repository.Update(entity);
                }
                return res;
            }
            else
            {
                return base.Delete(entity);
            }
        }

        public JsonResultModel<Atd_LeaveForm> GoProcess(Atd_LeaveForm entity)
        {
            JsonResultModel<Atd_LeaveForm> result = new JsonResultModel<Atd_LeaveForm>();
            var settingEntity = settingRepository.Entities.FirstOrDefault(i => i.LeaveType == entity.LeaveType);
            if (settingEntity == null)
            {
                result.Code = "100";
                result.Result = null;
                result.Message = "未设置假期类型，无法请假";
                return result;
            }
            var year = DateTime.Now.Year;
            var balanceEntity = balanceRepository.Entities.Where(i => i.UserId == entity.UserId && i.Year == year);
            var annualEntity = balanceRepository.Entities.FirstOrDefault(i => i.LeaveType == "annual" && i.UserId == entity.UserId && i.Year == year);
            var sickEntity = balanceRepository.Entities.FirstOrDefault(i => i.LeaveType == "sick" && i.UserId == entity.UserId && i.Year == year);
            var personalEntity = balanceRepository.Entities.FirstOrDefault(i => i.LeaveType == "personal" && i.UserId == entity.UserId && i.Year == year);



            if (entity.LeaveType == "annual")
            {
                if (annualEntity == null)
                {
                    result.Code = "100";
                    result.Result = null;
                    result.Message = "年假数据为空，无法计算，请假失败";
                    return result;
                }
                else
                {
                    if (annualEntity.Available < entity.TotalHours)
                    {
                        result.Code = "100";
                        result.Result = null;
                        result.Message = $"请假{entity.TotalHours}h 超过年假可用{annualEntity.Available} h，请假失败";
                        return result;
                    }
                    else
                    {
                        var res = InsertFormAndProcess(entity);
                        if (res != null)
                        {
                            OperateQuota(entity, balanceEntity, false, true);
                        }

                        result.Code = res == null ? "100" : "000";
                        result.Result = res;
                        result.Message = res == null ? "成功" : "失败";
                        return result;

                    }
                }
            }
            else if (entity.LeaveType == "personal")
            {
                if (annualEntity == null || personalEntity == null)
                {
                    result.Code = "100";
                    result.Result = null;
                    result.Message = "假期数据为空，无法计算，请假失败";
                    return result;
                }
                else
                {
                    if (annualEntity.Available != 0)
                    {
                        result.Code = "100";
                        result.Result = null;
                        result.Message = $"请假失败！当前还有{annualEntity.Available} h年假未使用完，请优先使用年假。";
                        return result;
                    }
                    else
                    {
                        if (personalEntity.Available < entity.TotalHours)
                        {
                            result.Code = "100";
                            result.Result = null;
                            result.Message = $"请假{entity.TotalHours}h 超过每年可用事假额度，请假失败";
                            return result;
                        }
                        var res = InsertFormAndProcess(entity);
                        if (res != null)
                        {
                            OperateQuota(entity, balanceEntity, false, true);
                        }
                        result.Code = res == null ? "100" : "000";
                        result.Result = res;
                        result.Message = res == null ? "成功" : "失败";
                        return result;
                    }
                }
            }
            else if (entity.LeaveType == "parental")
            {
                var parentalEntity = balanceEntity.FirstOrDefault(i => i.LeaveType == entity.LeaveType);
                if (parentalEntity == null)
                {
                    result.Code = "100";
                    result.Result = null;
                    result.Message = "假期数据为空，无法计算，请假失败";
                    return result;
                }
                else
                {
                    if (parentalEntity.Available < entity.TotalHours)
                    {
                        result.Code = "100";
                        result.Result = null;
                        result.Message = $"请假{entity.TotalHours}h 超过每年可用额度，请假失败";
                        return result;
                    }
                    var res = InsertFormAndProcess(entity);
                    if (res != null)
                    {
                        OperateQuota(entity, balanceEntity, false, true);
                    }
                    result.Code = res == null ? "100" : "000";
                    result.Result = res;
                    result.Message = res == null ? "成功" : "失败";
                    return result;
                }
            }
            else
            {
                if (entity.LeaveType == "sick" && sickEntity != null)
                {

                    if (sickEntity.Available > entity.TotalHours)
                    {
                        var res = InsertFormAndProcess(entity);
                        if (res != null)
                        {
                            OperateQuota(entity, balanceEntity, false, true);
                        }
                        result.Code = res == null ? "100" : "000";
                        result.Result = res;
                        result.Message = res == null ? "成功" : "失败";
                        return result;
                    }
                    else
                    {
                        var res = InsertFormAndProcess(entity);
                        result.Code = res == null ? "100" : "000";
                        result.Result = res;
                        result.Message = res == null ? "成功" : "失败";
                        return result;
                    }

                }
                else
                {
                    var res = InsertFormAndProcess(entity);
                    result.Code = res == null ? "100" : "000";
                    result.Result = res;
                    result.Message = res == null ? "成功" : "失败";
                    return result;
                }

            }
        }

        public Atd_LeaveForm InsertFormAndProcess(Atd_LeaveForm entity)
        {
            Atd_LeaveForm resEntity = repository.Entities.FirstOrDefault(i => i.ID == entity.ID);
            if (resEntity == null)
            {
                resEntity = base.Insert(entity);
            }
            else
            {
                //resEntity =  entity.ToObject<Atd_LeaveForm>();
                base.Update(entity);
            }
            try
            {
                var res = processService.AddProcessByLeaveType(resEntity.LeaveId);
                if (res)
                {
                    //return resEntity;
                    return resEntity;
                }
                else
                {
                    repository.Delete(resEntity);
                    return null;
                }
            }
            catch (Exception ex)
            {
                logHelper.LogError("add process fail" + ex.ToString());
                repository.Delete(resEntity);
                return null;
            }
        }

        public JsonResultModel<Atd_LeaveForm> UpdateFP(Atd_LeaveForm entity)
        {
            var settingEntity = settingRepository.Entities.FirstOrDefault(i => i.LeaveType == entity.LeaveType);
            if (settingEntity == null)
                return new JsonResultModel<Atd_LeaveForm>()
                {
                    Code = "100",
                    Result = null,
                    Message = "fail"
                };

            var workTime = userRepository.FirstOrDefault(i => i.UserName == entity.UserId).Worktime;
            workTime = string.IsNullOrEmpty(workTime) ? "8:30-17:30" : workTime;
            var isContainHoliday = settingEntity.IsContainHoliday;
            var isNursing = settingEntity.LeaveType == "nursing";
            var totalHour = Convert.ToDecimal(CalculateLeaveHours(entity.StartDate, entity.EndDate, workTime, isContainHoliday, isNursing));

            entity.TotalHours = totalHour;
            //Convert.ToDecimal(CalculateTotalHours(entity.StartDate, entity.EndDate));

            if (entity.CurrentState == 0)
            {
                return GoProcess(entity);
            }
            else
            {
                var res = base.Update(entity);
                return new JsonResultModel<Atd_LeaveForm>()
                {
                    Code = res ? "000" : "100",
                    Result = entity,
                    Message = res ? "success" : "fail"
                };

            }
        }

        //public List<Atd_FormAndProcess> GetWaitAuditForm(PageRequest<Atd_LeaveForm> pageRequest, out int totalCount)
        //{
        //    var userId = pageRequest.defaultWhere.Split(',')[0].Split('=')[1];
        //    var user = userRepository.Entities.FirstOrDefault(i => i.UserName == userId);

        //    pageRequest.defaultWhere = null;

        //    var where = pageRequest.GetWhere().Compile();
        //    var processEntity = processRepository.Entities.Where(i => user.FullName.Contains(i.UserId) || user.UserName == i.UserId && i.Action == "audit").ToList();
        //    List<Atd_FormAndProcess> processForms = new List<Atd_FormAndProcess>();
        //    foreach (var process in processEntity)
        //    {
        //        var leaveId = process.LeaveId;
        //        var leaveProcess = processRepository.Entities.Where(i => i.LeaveId == leaveId && i.Action == "audit").OrderBy(i => i.OrderNo).ToList();
        //        var thisProcess = leaveProcess.FirstOrDefault(i => i.LeaveId == leaveId && user.FullName.Contains(i.UserId) || user.UserName == i.UserId);
        //        var index = leaveProcess.IndexOf(thisProcess);
        //        bool preProcessIsSuceess = true;
        //        if (index > 0)
        //        {
        //            preProcessIsSuceess = leaveProcess[index - 1].ProcessState == "success";
        //        }


        //        var forms = repository.Entities.Where(where).ToList();
        //        var form = forms.FirstOrDefault(i => i.LeaveId == leaveId && process.ProcessState == "wait" && preProcessIsSuceess);
        //        var processForm = form.ToObject<Atd_FormAndProcess>();
        //        if (form != null)
        //        {
        //            processForm.ProcessID = process.ID.ToString();
        //            processForm.FormID = form.ID.ToString();
        //            processForms.Add(processForm);
        //        }
        //    }
        //    totalCount = processForms.Count;
        //    var res = processForms.Skip((pageRequest.PageIndex - 1) * pageRequest.PageSize).Take(pageRequest.PageSize);
        //    return processForms;
        //}
        public bool ApproveByEmail(Atd_ApproveByEmail approveByEmail)
        {
            try
            {
                logHelper.LogInfo(approveByEmail.FromEmail+ "Approve By Email");
                var user = userRepository.FirstOrDefault(i => i.Email == approveByEmail.FromEmail);
                if (user == null)
                {
                    logHelper.LogError("ApproveByEmail User NotFound :" + approveByEmail.FromEmail);
                    return false;
                }

                PageRequest<Atd_LeaveForm> pageRequest = new PageRequest<Atd_LeaveForm>
                {
                    defaultWhere = $"userId={user.UserName}",
                    order = "createDate",
                    PageIndex = 1,
                    PageSize = 10,
                    where = $"leaveId={approveByEmail.LeaveId}"
                };
                var waitAuditRes = GetWaitAuditForm(pageRequest, out int total).FirstOrDefault();
                if (waitAuditRes == null)
                {
                    var entity = processRepository.FirstOrDefault(i => i.LeaveId == approveByEmail.LeaveId && (i.UserId == user.UserName || user.FullName.Contains(i.UserId)));
                    if (entity != null)
                    {
                        return entity.ProcessState=="success"||entity.ProcessState=="error";
                    }
                    logHelper.LogError("ApproveByEmail Form NotFound :" + approveByEmail.LeaveId);
                    return false;
                }
                Atd_Audit auditRequest = new()
                {
                    ProcessID = waitAuditRes.ProcessID,
                    FormID = waitAuditRes.FormID,
                    Comment = approveByEmail.Comment,
                    Result = approveByEmail.Result,
                };
                var auditRes = AuditForm(auditRequest);
                return auditRes;
            }
            catch (Exception ex)
            {
                logHelper.LogError("ApproveByEmail Exception :" + ex.ToString());
                return false;
            }
        }

        public List<Atd_FormAndProcess> GetWaitAuditForm(PageRequest<Atd_LeaveForm> pageRequest, out int totalCount)
        {
            var userId = pageRequest.defaultWhere.Split(',')[0].Split('=')[1];
            var user = userRepository.Entities.FirstOrDefault(i => i.UserName == userId);

            pageRequest.defaultWhere = null;

            var where = pageRequest.GetWhere().Compile();
            //var order = pageRequest.GetOrder().Compile();

            var processEntity = processRepository.Entities.Where(i => user.FullName.Contains(i.UserId) || user.UserName == i.UserId && i.Action == "audit").ToList();
            List<Atd_FormAndProcess> processForms = new List<Atd_FormAndProcess>();

            var leaveIds = processEntity.Select(p => p.LeaveId).ToList(); // 提取所有 leaveId

            // 获取所有符合条件的 forms，一次性查询
            var forms = repository.Entities.Where(where).ToList();

            foreach (var process in processEntity)
            {
                var leaveId = process.LeaveId;
                var leaveProcess = processRepository.Entities.Where(i => i.LeaveId == leaveId && i.Action == "audit").OrderBy(i => i.OrderNo).ToList();
                var thisProcess = leaveProcess.FirstOrDefault(i => i.LeaveId == leaveId && user.FullName.Contains(i.UserId) || user.UserName == i.UserId);
                var index = leaveProcess.IndexOf(thisProcess);
                bool preProcessIsSuccess = true;
                if (index > 0)
                {
                    preProcessIsSuccess = leaveProcess[index - 1].ProcessState == "success";
                }

                // 使用之前提取的 forms 列表进行查询
                var form = forms.FirstOrDefault(i => i.LeaveId == leaveId && process.ProcessState == "wait" && preProcessIsSuccess);
                if (form != null)
                {
                    form.Attachment = string.IsNullOrEmpty(form.Attachment) ? "no" : "yes";
                    var processForm = form.ToObject<Atd_FormAndProcess>();
                    processForm.ProcessID = process.ID.ToString();
                    processForm.FormID = form.ID.ToString();
                    processForms.Add(processForm);
                }
            }

            totalCount = processForms.Count;
            //var res = new List<Atd_FormAndProcess>();

            //if (!pageRequest.isDesc)
            //{
            //    res = processForms.OrderBy(order).Skip((pageRequest.PageIndex - 1) * pageRequest.PageSize).Take(pageRequest.PageSize);
            //}
            //else
            //{
            //    res = processForms.OrderByDescending(order).Skip((pageRequest.PageIndex - 1) * pageRequest.PageSize).Take(pageRequest.PageSize);
            //}

            var res = processForms.Skip((pageRequest.PageIndex - 1) * pageRequest.PageSize).Take(pageRequest.PageSize).ToList();
            return res;
        }

        public bool AuditForm(Atd_Audit T)
        {
            var processEntity = processRepository.GetById(Guid.Parse(T.ProcessID));
            processEntity.Result = T.Comment;
            processEntity.AuditTime = DateTime.Now;
            processEntity.ProcessState = T.Result == "agree" ? "success" : "error";

            var formEntity = repository.GetById(Guid.Parse(T.FormID));
            var year = formEntity.StartDate.Year;

            var balanceEntity = balanceRepository.Entities.Where(i => i.UserId == formEntity.UserId && i.Year == year);

            var user = userRepository.Entities.FirstOrDefault(i => i.UserName == formEntity.UserId);
            var auditUser = userRepository.Entities.FirstOrDefault(i => i.UserName == processEntity.UserId || i.FullName.Contains(processEntity.UserId));

            if (T.Result == "deny")
            {
                formEntity.CurrentState = 3;
                OperateQuota(formEntity, balanceEntity, false);
                SendMailByCurrentState(user, auditUser, null, formEntity);
            }
            else
            {
                if (processEntity.IsLastNode)
                {
                    formEntity.CurrentState = 2;
                    if (!formEntity.IsCancel)
                    {
                        formEntity.CancelState = 2;
                    }
                    OperateQuota(formEntity, balanceEntity, true);
                    SendMailByCurrentState(user, auditUser, null, formEntity);
                }
                else
                {
                    formEntity.CurrentState = 1;
                    var nextProcessEntity = processRepository.FirstOrDefault(i => i.LeaveId == formEntity.LeaveId && i.OrderNo == processEntity.OrderNo + 1);
                    var nextUser = userRepository.Entities.FirstOrDefault(i => i.UserName == nextProcessEntity.UserId || i.FullName.Contains(nextProcessEntity.UserId));

                    SendMailByCurrentState(user, auditUser, nextUser, formEntity);
                }
            }

            try
            {
                //processRepository.BeginTransaction();

                processRepository.Update(processEntity);
                repository.Update(formEntity);

                //processRepository.Commit();
                return true;
            }
            catch (Exception ex)
            {
                logHelper.LogError("audit error" + ex.ToString());
                processRepository.Rollback();
                return false;
            }

        }

        //public bool ForwardTo(string toEmail,string formId)
        //{
        //    var form = repository.FirstOrDefault(i => i.ID.ToString() == formId);
        //    var url = configuration.GetSection("EmailNotify:Url").Value;
        //    string fromEmail = auditUser.Email;
        //    string title = "";
        //    string body = "";
        //    title = $"Time off Request from {user.FullName}";
        //    body = $"Dear {auditUser.FullName} ,<br />";
        //    body += $"A leave request {form.LeaveId} requires your approval. <br />";
        //    body += $"Please go to {url}/leaveApprove?leaveId={fEntity.LeaveId} to view details."; 
        //}

        public void SendMailByCurrentState(Sys_Users user, Sys_Users auditUser, Sys_Users nextUser, Atd_LeaveForm fEntity)
        {
            var url = configuration.GetSection("EmailNotify:Url").Value;
            string fromEmail = auditUser.Email;
            string toEmail = user.Email;

            string title = "";
            string body = "";

            //string leaveDetail = $"{user.FullName}<br/>";
            //leaveDetail += $"{fEntity.LeaveType} Vacation {fEntity.TotalHours} hours {GetHour2Day(fEntity.TotalHours)}<br/>";
            //leaveDetail += $"From {fEntity.StartDate:yyyy-MM-dd HH:mm} To {fEntity.EndDate:yyyy-MM-dd HH:mm} <br/>";
            //leaveDetail += $"Reason : {fEntity.Reason ?? ""}<br/>";
            string leaveDetail = @"<div><table width='100%' border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse;'>"; 
            leaveDetail += @"<tr>
            <th colspan='1' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>Employee</th>
            <th colspan='1' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>Request</th>
            <th colspan='2' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>From</th>
            <th colspan='2' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>To</th>
            <th colspan='1' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>Total Days</th>
            <th colspan='1' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>Total Hours</th>
            </tr>";
            leaveDetail += "<tr>" +
                $"<td colspan='1' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{user.FullName}</td>" +
                $"<td colspan='1' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.LeaveType} Leave</td>" +
                $"<td colspan='2' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.StartDate:yyyy-MM-dd HH:mm}</td>" +
                $"<td colspan='2' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.EndDate:yyyy-MM-dd HH:mm}</td>" +
                $"<td colspan='1' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{(fEntity.TotalHours / 8).Value.ToString("0.00")}</td>" +
                $"<td colspan='1' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.TotalHours}</td>" +
                "</tr>";
            leaveDetail += @"<tr><td>  &nbsp  <br/></td></tr>";
            leaveDetail += @" <tr>
            <th colspan='4' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>Reason</th> 
            </tr>";
            leaveDetail += $"<tr><td colspan='4' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.Reason??""}</td></tr>"; 
            leaveDetail += "</table></div><br/>";


            if (fEntity.CurrentState == 0)
            {
                title = $"Time off Request from {user.FullName}"; 

                body = $"<span style='font-family:Arial;font-size:11pt;'>Dear </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{auditUser.FullName}</span> ,<br/>";
                body += $"<span style='font-family:Arial;font-size:11pt;'>A leave request: </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.LeaveId}</span><span style='font-family:Arial;font-size:11pt;'> requires your approval: </span><br/><br/><br/>";

                body += leaveDetail;

                body += "<div width='100%' ><span style='color: black;font-weight:bold;font-family:Arial;font-size:11pt;'>Action</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> ( Please click below links to </span><span style='color:#808080;font-weight:bold;font-style:italic;font-family:Arial;font-size:10pt;'>Approve</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> or </span><span style='color:#808080;font-weight:bold;font-style:italic;font-family:Arial;font-size:10pt;'>Reject</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> the leave request by replying with email directly )</span></div> <br/>";

                body += @$"<div width='100%'>
<span><a href='mailto:HRVac@ehealth.com?subject=leaveId={fEntity.LeaveId}-result=agree&body=Comment:'>Approve</a></span>
&nbsp | &nbsp
<span><a href='mailto:HRVac@ehealth.com?subject=leaveId={fEntity.LeaveId}-result=deny&body=Comment:'>Reject</a></span>
</div> <br/>";

                body += @$"<div width='100%'>
<span><a href='{url}/leaveApprove?leaveId={fEntity.LeaveId}'>View details</a></span>
<span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> ( Log in HR vacation system to review request details )</span>
</div> <br/>"; 

                toEmail = auditUser.Email;
            }

            if (fEntity.CurrentState == 1)
            {
                title = $"Time off Request from {user.FullName}";
                body = $"<span style='font-family:Arial;font-size:11pt;'>Dear </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{nextUser.FullName}</span> ,<br/>";
                body += $"<span style='font-family:Arial;font-size:11pt;'>A leave request: </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.LeaveId}</span><span style='font-family:Arial;font-size:11pt;'> requires your approval: </span><br/><br/><br/>";

                body += leaveDetail;

                body += "<div width='100%' ><span style='color: black;font-weight:bold;font-family:Arial;font-size:11pt;'>Action</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> ( Please click below links to </span><span style='color:#808080;font-weight:bold;font-style:italic;font-family:Arial;font-size:10pt;'>Approve</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> or </span><span style='color:#808080;font-weight:bold;font-style:italic;font-family:Arial;font-size:10pt;'>Reject</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> the leave request by replying with email directly )</span></div> <br/>";

                body += @$"<div width='100%'>
<span><a href='mailto:HRVac@ehealth.com?subject=leaveId={fEntity.LeaveId}-result=agree&body=Comment:'>Approve</a></span>
&nbsp | &nbsp
<span><a href='mailto:HRVac@ehealth.com?subject=leaveId={fEntity.LeaveId}-result=deny&body=Comment:'>Reject</a></span>
</div> <br/>";

                body += @$"<div width='100%'>
<span><a href='{url}/leaveApprove?leaveId={fEntity.LeaveId}'>View details</a></span>
<span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> ( Log in HR vacation system to review request details )</span>
</div> <br/>";
                toEmail = nextUser.Email;
            }
            if (fEntity.CurrentState == 2)
            {
                title = "Your leave request has been approved";
                body = $"<span style='font-family:Arial;font-size:11pt;'>Dear </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{user.FullName}</span> ,<br/>";
                body += $"<span style='font-family:Arial;font-size:11pt;'>Your leave request: </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.LeaveId}</span><span style='font-family:Arial;font-size:11pt;'> has been approved by  </span><br/><br/><br/>";
                body += $"Your leave request: {fEntity.LeaveId} has been approved by <span style='color: #000099;font-family:Arial;font-size:11pt;'>{auditUser.FullName}</span> <br/>";

                body += leaveDetail;

                body += @$"<div width='100%'>
<span><a href=""{url}/myLeave?leaveId={fEntity.LeaveId}"">View details</a></span>
<span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> ( Log in HR vacation system to review request details )</span>
</div> <br/>";
                 
            }
            if (fEntity.CurrentState == 3)
            {
                title = "Your leave request has been rejected";
                body = $"<span style='font-family:Arial;font-size:11pt;'>Dear </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{user.FullName}</span> ,<br/>";
                body += $"<span style='font-family:Arial;font-size:11pt;'>Sorry, your leave request: </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{fEntity.LeaveId}</span><span style='font-family:Arial;font-size:11pt;'> has been rejected.<br/>Please reapply  </span><br/><br/><br/>";
                //body += "<span style='font-family:Arial;font-size:11pt;'>Please reapply</span>";
                body += leaveDetail;
             
                body += @$"<div width='100%'>
<span><a href=""{url}/myLeave?leaveId={fEntity.LeaveId}"">View details</a></span>
<span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> ( Log in HR vacation system to review request details )</span>
</div> <br/>";
            }

            var needSend = Convert.ToBoolean(configuration.GetSection("EmailNotify:IsRequire").Value);

            if (needSend)
            {
                List<string> CC = new List<string>();
                if (!string.IsNullOrEmpty(user.CC))
                {
                    if (user.CCHours != null)
                    {
                        if (fEntity.TotalHours >= user.CCHours)
                        {
                            CC.Add(user.CC);
                        }
                    }
                }
                if (toEmail == nextUser?.Email || toEmail == auditUser?.Email)
                {
                    
                }
                else
                {
                    CC.Add(user.Email);
                }
                mailService.Send(fromEmail, toEmail, title, body, CC);
            }

        }

        public string GetHour2Day(decimal? totalHour)
        {
            int days = (int)(totalHour / 8); // 计算天数
            var hour = totalHour % 8;
            var result = $"({days} day {hour} h)";
            return result;
        }

        public void OperateQuota(Atd_LeaveForm fEntity, IQueryable<Atd_LeaveBalance> balanceEntity, bool isAgree, bool isApply = false)
        {
            var bEntity = balanceEntity.FirstOrDefault(i => i.LeaveType == fEntity.LeaveType);
            if (bEntity != null)
            {
                if (fEntity.IsCancel)
                {
                    bEntity.Available += (decimal)fEntity.TotalHours;
                    if (fEntity.CancelState == 2)
                    {
                        bEntity.Used -= (decimal)fEntity.TotalHours;
                    }
                    else
                    {
                        bEntity.Locked -= (decimal)fEntity.TotalHours;
                    }
                }
                else
                {
                    if (isApply)
                    {
                        bEntity.Locked += (decimal)fEntity.TotalHours;
                        bEntity.Available -= (decimal)fEntity.TotalHours;
                    }
                    else
                    {
                        if (!isAgree)
                        {

                            bEntity.Locked -= (decimal)fEntity.TotalHours;
                            bEntity.Available += (decimal)fEntity.TotalHours;
                        }
                        else
                        {
                            bEntity.Locked -= (decimal)fEntity.TotalHours;
                            bEntity.Used += (decimal)fEntity.TotalHours;
                        }
                    }
                }

                balanceRepository.Update(bEntity);
            }

        }

        public bool SendMail(string formEmail, string toEmail, string subject, string body)
        {
            mailService.Send(formEmail, toEmail, subject, body);
            return true;
        }

        public double CalculateLeaveHours(DateTime dtStart, DateTime dtEnd, string workTime, bool isContainHoliday, bool isNursing = false)
        {
            var startWorkStr = workTime.Split("-")[0].Split(":");
            var startWorkHour = Convert.ToInt32(startWorkStr[0]);
            var startWorkMinute = Convert.ToInt32(startWorkStr[1]);
            var endWorkStr = workTime.Split("-")[1].Split(":");
            var endWorkHour = Convert.ToInt32(endWorkStr[0]);
            var endWorkMinute = Convert.ToInt32(endWorkStr[1]);

            double leaveDuration = 0;
            if (isContainHoliday)
            {
                leaveDuration = ((dtEnd.Date - dtStart.Date).Days + 1) * 8;
            }
            else
            {
                DateTime startWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, startWorkHour, startWorkMinute, 0);//上班时间 

                dtStart = dtStart < startWork ? startWork : dtStart;

                ////上下班时间 
                //TimeSpan workStartTime = new TimeSpan(startWorkHour, startWorkMinute, 0);
                //TimeSpan workEndTime = new TimeSpan(endWorkHour, endWorkMinute, 0); 
                //// 中午休息时间范围（12:00 ~ 13:00）
                //TimeSpan lunchStartTime = new TimeSpan(12, 0, 0);
                //TimeSpan lunchEndTime = new TimeSpan(13, 0, 0);

                // 遍历从开始时间到结束时间的每一天
                for (DateTime date = dtStart.Date; date <= dtEnd.Date; date = date.AddDays(1))
                {
                    // 上午请假时间
                    DateTime morningStart = date.Date.AddHours(startWorkHour).AddMinutes(startWorkMinute);
                    DateTime morningEnd = date.Date.AddHours(12);

                    // 下午请假时间
                    DateTime afternoonStart = date.Date.AddHours(13);
                    DateTime afternoonEnd = date.Date.AddHours(endWorkHour).AddMinutes(endWorkMinute);
                    // 排除周末和假日
                    if (IsWorkDay(date, null))
                    {
                        if (isNursing)
                        {
                            leaveDuration += 1;
                            continue;
                        }
                        // 如果请假开始时间在下班之后，则当天不算请假时长
                        if (dtStart >= afternoonEnd)
                        {
                            dtStart = dtStart.Date.AddDays(1).AddHours(startWorkHour).AddMinutes(startWorkMinute);
                            continue;
                        }
                        // 如果请假结束时间在上班之前，则当天不算请假时长
                        if (dtEnd <= morningStart)
                        {
                            dtEnd = dtEnd.Date.AddDays(-1).AddHours(endWorkHour).AddMinutes(endWorkMinute);
                            continue;
                        }

                        // 调整请假开始时间和结束时间，确保在工作时间内
                        DateTime adjustedStartTime = dtStart > morningStart ? dtStart : morningStart;
                        adjustedStartTime = (adjustedStartTime > morningEnd && adjustedStartTime < afternoonStart) ? morningEnd : adjustedStartTime;
                        DateTime adjustedEndTime = dtEnd < afternoonEnd ? dtEnd : afternoonEnd;

                        // 计算每天的请假时长
                        double dailyLeaveDuration = (adjustedEndTime - adjustedStartTime).TotalHours;
                        if (adjustedEndTime >= afternoonStart && adjustedStartTime <= morningEnd)
                            dailyLeaveDuration -= 1;


                        // 累加每天的请假时长
                        leaveDuration += dailyLeaveDuration;
                    }
                }
            }

            return RoundToNearestHalf(leaveDuration);
        }

        public double CalculateTotalHours(DateTime dtStart, DateTime dtEnd, Dictionary<string, object> workDays = null)
        {

            DateTime dtFirstDayGoToWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 8, 30, 0);//请假第一天的上班时间
            DateTime dtFirstDayGoOffWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 17, 30, 0);//请假第一天的下班时间

            DateTime dtLastDayGoToWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 8, 30, 0);//请假最后一天的上班时间
            DateTime dtLastDayGoOffWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 17, 30, 0);//请假最后一天的下班时间

            DateTime dtFirstDayRestStart = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 12, 00, 0);//请假第一天的午休开始时间
            DateTime dtFirstDayRestEnd = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 13, 00, 0);//请假第一天的午休结束时间

            DateTime dtLastDayRestStart = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 12, 00, 0);//请假最后一天的午休开始时间
            DateTime dtLastDayRestEnd = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 13, 00, 0);//请假最后一天的午休结束时间

            //如果开始请假时间早于上班时间或者结束请假时间晚于下班时间，者需要重置时间
            //if (!IsWorkDay(dtStart, workDays) && !IsWorkDay(dtEnd, workDays))
            //    return 0;
            //if (dtStart >= dtFirstDayGoOffWork && dtEnd <= dtLastDayGoToWork && (dtEnd - dtStart).TotalDays < 1)
            //    return 0;
            //if (dtStart >= dtFirstDayGoOffWork && !IsWorkDay(dtEnd, workDays) && (dtEnd - dtStart).TotalDays < 1)
            //    return 0;

            if (dtEnd < dtStart)
                return 0;

            var tempStartDate = dtStart;
            var tempEndDate = dtEnd;

            if (dtStart < dtFirstDayGoToWork)//早于上班时间
                dtStart = dtFirstDayGoToWork;
            if (dtStart >= dtFirstDayGoOffWork)//晚于下班时间
            {
                while (tempStartDate < tempEndDate)
                {
                    dtStart = new DateTime(dtStart.AddDays(1).Year, dtStart.AddDays(1).Month, dtStart.AddDays(1).Day, 8, 30, 0);
                    if (IsWorkDay(dtStart, workDays))
                    {
                        dtFirstDayGoToWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 8, 30, 0);//请假第一天的上班时间
                        dtFirstDayGoOffWork = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 17, 30, 0);//请假第一天的下班时间
                        dtFirstDayRestStart = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 11, 30, 0);//请假第一天的午休开始时间
                        dtFirstDayRestEnd = new DateTime(dtStart.Year, dtStart.Month, dtStart.Day, 12, 30, 0);//请假第一天的午休结束时间

                        break;
                    }
                }
            }

            if (dtEnd > dtLastDayGoOffWork)//晚于下班时间
                dtEnd = dtLastDayGoOffWork;
            if (dtEnd <= dtLastDayGoToWork)//早于上班时间
            {
                while (tempEndDate > tempStartDate)
                {
                    dtEnd = new DateTime(dtEnd.AddDays(-1).Year, dtEnd.AddDays(-1).Month, dtEnd.AddDays(-1).Day, 17, 30, 0);
                    if (IsWorkDay(dtEnd, workDays))//
                    {
                        dtLastDayGoToWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 8, 30, 0);//请假最后一天的上班时间
                        dtLastDayGoOffWork = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 17, 30, 0);//请假最后一天的下班时间
                        dtLastDayRestStart = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 11, 30, 0);//请假最后一天的午休开始时间
                        dtLastDayRestEnd = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day, 12, 30, 0);//请假最后一天的午休结束时间
                        break;
                    }
                }
            }

            //计算请假第一天和最后一天的小时合计数并换算成分钟数           
            double iSumMinute = dtFirstDayGoOffWork.Subtract(dtStart).TotalMinutes + dtEnd.Subtract(dtLastDayGoToWork).TotalMinutes;//计算获得剩余的分钟数

            if (dtStart >= dtFirstDayRestStart && dtStart <= dtFirstDayRestEnd)
            {//开始休假时间正好是在午休时间内的，需要扣除掉
                iSumMinute -= dtFirstDayRestEnd.Subtract(dtStart).Minutes;
            }
            if (dtStart <= dtFirstDayRestStart)
            {//如果是在午休前开始休假的就自动减去午休的60分钟
                iSumMinute -= 60;
            }
            if (dtEnd >= dtLastDayRestStart && dtEnd <= dtLastDayRestEnd)
            {//如果结束休假是在午休时间内的，例如“请假截止日是1月31日 12:00分”的话那休假时间计算只到 11:30分为止。
                iSumMinute -= dtEnd.Subtract(dtLastDayRestStart).Minutes;
            }
            if (dtEnd >= dtLastDayRestEnd)
            {//如果是在午休后结束请假的就自动减去午休的60分钟
                iSumMinute -= 60;
            }


            int leaveday = 0;//实际请假的天数
            double countday = 0;//获取两个日期间的总天数

            DateTime tempDate = dtStart;//临时参数
            while (tempDate < dtEnd)
            {
                countday++;
                tempDate = new DateTime(tempDate.AddDays(1).Year, tempDate.AddDays(1).Month, tempDate.AddDays(1).Day, 0, 0, 0);
            }
            //循环用来扣除双休日、法定假日 和 添加调休上班
            for (int i = 0; i < countday; i++)
            {
                DateTime tempdt = dtStart.Date.AddDays(i);

                if (IsWorkDay(tempdt, workDays))
                    leaveday++;
            }

            //去掉请假第一天和请假的最后一天，其余时间全部已8小时计算。
            //SumMinute/60： 独立计算 请假第一天和请假最后一天总归请了多少小时的假
            double doubleSumHours = RoundToNearestHalf(((leaveday - 2) * 8) + iSumMinute / 60);
            //int intSumHours = Convert.ToInt32(doubleSumHours);

            //if (doubleSumHours > intSumHours)//如果请假时间不足1小时话自动算作1小时
            //    intSumHours++;

            return doubleSumHours;

        }

        private bool IsWorkDay(DateTime date, Dictionary<string, object> workDays)
        {
            try
            {
                var m_simVacationData = holidayService.GetHoliday(date.Year);
                //读取数据库中【Vacation】表中的所有数据,返回一个Datatable等等...
                //我这里采用的是内存操作Dictionary，因为一般这种节假日都是固定不变的，不需要每次都取访问数据查询一遍。

                //Dictionary<string, Vacation> m_simVacationData = newDictionary<string, SimVacation>(); 

                //利用Datatable的值循环给 m_simVacationData 赋值。

                string DateKey = date.ToString("yyyy-MM-dd");//日期值：“2012-08-01”

                bool b_wokrdate = !m_simVacationData.Contains(DateKey);
                return b_wokrdate;

            }
            catch (Exception ex)
            {
                //抛出异常
                logHelper.LogError("IsWorkDay error" + ex.ToString());
                return false;
            }


            ////星期天并且不属于节假日和调休上班
            //if (date.DayOfWeek == DayOfWeek.Sunday && !m_simVacationData.Contains(DateKey))
            //    return false;
            ////星期六并且不属于节假日和调休上班     

            //else if (date.DayOfWeek == DayOfWeek.Saturday && !m_simVacationData.Contains(DateKey))

            //    return false;
            //else if (m_simVacationData.Contains(DateKey))//属于节假日或调休
            //{
            //    if (m_simVacationData[DateKey].Is_Vacation)//Is_Vacation=true（节假日） Is_Vacation=false（调休上班）
            //    {
            //        return false;
            //    }
            //}
            //if (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday)
            //{
            //    b_wokrdate = false;
            //}


        }

        public static double RoundToNearestHalf(double value)
        {
            int rounded = (int)Math.Round(value / 0.5);
            return rounded * 0.5;
        }
    }

}
