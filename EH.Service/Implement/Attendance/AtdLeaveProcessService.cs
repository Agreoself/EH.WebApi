using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Interface.Attendance;
using EH.Service.Interface.Attendance;
using EH.Repository.Interface.Sys;
using Microsoft.Extensions.Options;
using Microsoft.VisualBasic.FileIO;
using NPOI.HSSF.Record.Chart;
using EH.System.Models.Dtos;
using Microsoft.Extensions.Configuration;
using Microsoft.EntityFrameworkCore;

namespace EH.Service.Implement.Attendance
{
    public class AtdLeaveProcessService : BaseService<Atd_LeaveProcess>, IAtdLeaveProcessService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IAtdLeaveProcessRepository repository;
        private readonly IAtdLeaveFormRepository formRepository;
        private readonly IAtdLeaveSettingRepository settingRepository;
        private readonly ISysUsersRepository userRepository;

        private readonly IMailService mailService;
        private readonly IConfiguration configuration;

        public AtdLeaveProcessService(IAtdLeaveProcessRepository repository, IAtdLeaveFormRepository formRepository,
            IAtdLeaveSettingRepository settingRepository, ISysUsersRepository userRepository, IMailService mailService, IConfiguration configuration, LogHelper logHelper) : base(repository, logHelper)
        {
            this.logHelper = logHelper;
            this.repository = repository;
            this.formRepository = formRepository;
            this.settingRepository = settingRepository;
            this.userRepository = userRepository;
            this.mailService = mailService;
            this.configuration = configuration;
        }

        public override List<Atd_LeaveProcess> GetPageList(PageRequest<Atd_LeaveProcess> request, out int totalCount)
        {
            var res = base.GetPageList(request, out totalCount);
            foreach (var item in res)
            {
                //item.UserId = userRepository.FirstOrDefault(i => i.UserName == item.UserId).FullName;
                item.UserId = userRepository.FirstOrDefault(i => i.UserName == item.UserId || i.FullName.Contains(item.UserId)).FullName;
            }
            return res;
        }

        public bool AddProcessByLeaveType(string leaveId)
        {
            var formEntity = formRepository.Entities.FirstOrDefault(i => i.LeaveId == leaveId);
            if (formEntity == null)
            {
                logHelper.LogError("请假数据未找到");
                return false;
            }
            else
            {
                var hrName = configuration.GetSection("EmailNotify:HrName").Value;
                var hrEmail = configuration.GetSection("EmailNotify:HrEmail").Value;

                var hr = userRepository.FirstOrDefault(i => i.UserName == hrName);
                var user = userRepository.FirstOrDefault(i => i.UserName == formEntity.UserId);
                //var lastReport = new string[] { "Laker Zeng", "lakerz" };
                //var reportIsLaker = lastReport.Contains(user.Report);
                var reportUser = userRepository.FirstOrDefault(i => i.UserName == user.Report || i.FullName.Contains(user.Report));

                var lastReportUser = userRepository.FirstOrDefault(i => i.UserName == user.LastReport || i.FullName.Contains(user.LastReport));

                var settingEntity = settingRepository.Entities.FirstOrDefault(i => i.LeaveType == formEntity.LeaveType);
                if (settingEntity == null)
                {
                    logHelper.LogError("假期未设置");
                    return false;
                }

                var needVpHour = settingEntity.NeedVpHour;//获得需要vp审核的请假小时数
                var needHrAudit = settingEntity.NeedHRApprove;

                List<Atd_LeaveProcess> list = new List<Atd_LeaveProcess>();

                Atd_LeaveProcess applyProcess = new()
                {
                    ProcessId = DateTime.Now.ToString("yyyyMMdd") + new Random().Next(1000),
                    LeaveId = formEntity.LeaveId,
                    UserId = formEntity.UserId,
                    Action = "apply",
                    Result = "Success",
                    AuditTime = DateTime.Now,
                    OrderNo = 1,
                    ProcessState = "success",
                    IsLastNode = false,
                };


                Atd_LeaveProcess auditProcess = new()
                {
                    ProcessId = DateTime.Now.ToString("yyyyMMdd") + new Random().Next(1000),
                    LeaveId = formEntity.LeaveId,
                    UserId = user.Report,
                    Action = "audit",
                    Result = "",
                    AuditTime = null,
                    OrderNo = 2,
                    ProcessState = "wait",
                };


                Atd_LeaveProcess vpAuditProcess = new()
                {
                    ProcessId = DateTime.Now.ToString("yyyyMMdd") + new Random().Next(1000),
                    LeaveId = formEntity.LeaveId,
                    UserId = user.LastReport,
                    Action = "audit",
                    Result = "",
                    AuditTime = null,
                    OrderNo = 3,
                    ProcessState = "wait",
                };

                Atd_LeaveProcess hrAuditProcess = new()
                {
                    ProcessId = DateTime.Now.ToString("yyyyMMdd") + new Random().Next(1000),
                    LeaveId = formEntity.LeaveId,
                    UserId = hr.UserName,
                    Action = "audit",
                    Result = "",
                    AuditTime = null,
                    OrderNo = 4,
                    ProcessState = "wait",
                    IsLastNode = true,
                };

                if (formEntity.IsCancel)//如果是销假，则找到原来的数据，在最后的数据后面加入两条流程
                {
                    try
                    {
                        var processEntitys = repository.Where(i => i.LeaveId == formEntity.LeaveId).OrderBy(i=>i.OrderNo);
                        var lastEntity = processEntitys.FirstOrDefault(i => i.ProcessState == "wait");
                        var orderNo = lastEntity == null ? processEntitys.Count() : lastEntity.OrderNo;
                        repository.DeleteRange(processEntitys.Where(i => i.OrderNo >= orderNo));

                        Atd_LeaveProcess cancelApplyProcess = new()
                        {
                            ProcessId = DateTime.Now.ToString("yyyyMMdd") + new Random().Next(1000),
                            LeaveId = formEntity.LeaveId,
                            UserId = formEntity.UserId,
                            Action = "cancel apply",
                            Result = "Success",
                            AuditTime = DateTime.Now,
                            OrderNo = orderNo,
                            ProcessState = "success",
                            IsLastNode = false,
                        };

                        Atd_LeaveProcess cancelHrAuditProcess = new()
                        {
                            ProcessId = DateTime.Now.ToString("yyyyMMdd") + new Random().Next(1000),
                            LeaveId = formEntity.LeaveId,
                            UserId = user.Report,
                            //UserId = hr.UserName,
                            Action = "audit",
                            Result = "",
                            AuditTime = null,
                            OrderNo = orderNo + 1,
                            ProcessState = "wait",
                            IsLastNode = true,
                        };
                        list.Add(cancelApplyProcess);
                        list.Add(cancelHrAuditProcess);

                    }
                    catch (Exception ex)
                    {
                        logHelper.LogError("cancelLeave fail: " + ex.ToString());
                        return false;
                    }

                }
                else
                {
                    list.Add(applyProcess);

                    var needVPApprove = formEntity.TotalHours >= needVpHour;

                    auditProcess.IsLastNode = (needVPApprove || needHrAudit) ? false : true;
                    list.Add(auditProcess);

                    if (needVPApprove)//如果大于就需要判断lastreport和report是不是同一个人，是则只需要report 批，不是则加一个vp批
                    {
                        if (reportUser.UserName != lastReportUser.UserName)
                        {
                            vpAuditProcess.IsLastNode = needHrAudit ? false : true;
                            list.Add(vpAuditProcess);
                        }
                        else
                        {
                            needVPApprove = false;
                            auditProcess.IsLastNode = needHrAudit ? false : true;
                        }
                    }

                    if (needHrAudit)
                    {
                        hrAuditProcess.OrderNo = needVPApprove ? 4 : 3;
                        list.Add(hrAuditProcess);
                    }
                }

                repository.AddRange(list);

                var url = configuration.GetSection("EmailNotify:Url").Value;
                //var report = formEntity.IsCancel ? "HR" : reportUser.FullName;
                string toEmail = configuration.GetSection("EmailNotify:SysAdminEmail").Value;

                string title = formEntity.IsCancel ? $"Request for cancellation of leave from {user.FullName}" : $"Time off Request from {user.FullName}";

               string body = $"<span style='font-family:Arial;font-size:11pt;'>Dear </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{reportUser.FullName}</span> ,<br/>";

                body += formEntity.IsCancel ? $"<span style='font-family:Arial;font-size:11pt;'>A cancellation leave request : </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{formEntity.LeaveId}</span><span style='font-family:Arial;font-size:11pt;'> requires your approval: </span><br/><br/><br/>": $"<span style='font-family:Arial;font-size:11pt;'>A leave request: </span><span style='color: #000099;font-family:Arial;font-size:11pt;'>{formEntity.LeaveId}</span><span style='font-family:Arial;font-size:11pt;'> requires your approval: </span><br/><br/><br/>";

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
                    $"<td colspan='1' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{formEntity.LeaveType} Leave</td>" +
                    $"<td colspan='2' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{formEntity.StartDate:yyyy-MM-dd HH:mm}</td>" +
                    $"<td colspan='2' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{formEntity.EndDate:yyyy-MM-dd HH:mm}</td>" +
                    $"<td colspan='1' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{(formEntity.TotalHours / 8).Value.ToString("0.00")}</td>" +
                    $"<td colspan='1' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{formEntity.TotalHours}</td>" +
                    "</tr>";
                leaveDetail += @"<tr><td>  &nbsp  <br/></td></tr>";
                leaveDetail += @" <tr>
            <th colspan='4' style='font-weight:bold;text-align: left;font-family:Arial;font-size:11pt;'>Reason</th> 
            </tr>";
                leaveDetail += $"<tr><td colspan='4' style='text-align: left;color: #000099;font-family:Arial;font-size:11pt;'>{formEntity.Reason ?? ""}</td></tr>";
                leaveDetail += "</table></div><br/>";
                body += leaveDetail;

                body += "<div width='100%' ><span style='color: black;font-weight:bold;font-family:Arial;font-size:11pt;'>Action</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> (<span style='font-style:italic;font-family:Arial;font-size:10pt;color:red'>For approvers only.</span>  Please click below links to </span><span style='color:#808080;font-weight:bold;font-style:italic;font-family:Arial;font-size:10pt;'>Approve</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> or </span><span style='color:#808080;font-weight:bold;font-style:italic;font-family:Arial;font-size:10pt;'>Reject</span><span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> the leave request by replying with email directly )</span></div> <br/>";

                body += @$"<div width='100%'>
<span><a href='mailto:HRVac@ehealth.com?subject=leaveId={formEntity.LeaveId}-result=agree&body=Comment:'>Approve</a></span>
&nbsp | &nbsp
<span><a href='mailto:HRVac@ehealth.com?subject=leaveId={formEntity.LeaveId}-result=deny&body=Comment:'>Reject</a></span>
</div> <br/>";

                body += @$"<div width='100%'>
<span><a href='{url}/leaveApprove?leaveId={formEntity.LeaveId}'>View details</a></span>
<span style='font-style:italic;font-family:Arial;font-size:10pt;color:#808080'> ( Log in HR vacation system to review request details )</span>
</div> <br/>";

                 
                string fromEmail = user.Email;
                var managerEntity = userRepository.Where(i => user.Report.Trim().Contains(i.UserName) || i.FullName.Contains(user.Report) || i.UserName == user.Report).FirstOrDefault();
                //logHelper.LogInfo("managerEntity:" + managerEntity == null ? "null" : managerEntity.ToJson());

                //toEmail = formEntity.IsCancel ? hrEmail : managerEntity == null ? toEmail : managerEntity.Email;
                toEmail = managerEntity == null ? toEmail : managerEntity.Email;

                var needSend = Convert.ToBoolean(configuration.GetSection("EmailNotify:IsRequire").Value);

                List<string> CC = new List<string>();
                CC.Add(user.Email);
                if (needSend)
                {
                    if (!string.IsNullOrEmpty(user.CC))
                    {
                        if (user.CCHours != null)
                        {
                            if (formEntity.TotalHours >= user.CCHours)
                            {
                                CC.Add(user.CC);
                            }
                        }
                    }
                     
                    mailService.Send(fromEmail, toEmail, title, body, CC);
                }

                return true;
            }
        }

        public string GetHour2Day(decimal? totalHour)
        {
            int days = (int)(totalHour / 8); // 计算天数
            var hour = totalHour % 8;
            var result = $"({days} day {hour} h)";
            return result;
        }

        public void GetUnapprovedUser()
        {
            List<NotifyUser> userList = new List<NotifyUser>();
            var allNeedProcess = repository.Entities.Where(i => i.Action == "audit" && i.ProcessState == "wait").ToList();
            foreach (var process in allNeedProcess)
            {
                var previous = repository.FirstOrDefault(i => i.LeaveId == process.LeaveId && i.OrderNo == process.OrderNo - 1);
                if (previous == null)
                    continue;
                if (previous.ProcessState == "success")
                {
                    var user = userRepository.FirstOrDefault(i => i.UserName == process.UserId || i.FullName.Contains(process.UserId));
                    if (user != null)
                    {
                        var userNotify = userList.FirstOrDefault(i => i.Users.ID == user.ID);
                        if (userNotify != null)
                        {
                            userNotify.count += 1;
                        }
                        else
                        {
                            NotifyUser nu = new NotifyUser();
                            nu.Users = user;
                            nu.count = 1;
                            userList.Add(nu);
                        }
                    }
                }
            } 
            userList.ForEach(i =>
            {
                logHelper.LogInfo($"{i.Users.FullName} have " + i.count+ " unapproved ");
                var url = configuration.GetSection("EmailNotify:Url").Value;
                string toEmail = i.Users.Email;
                string fromEmail = configuration.GetSection("Quartz:UnapprovedNotify:NotifyMail").Value;
                string title = $"There are leave requests that require your approval";
                string body = $"Dear {i.Users.FullName} ,<br />";
                body += $"You still have {i.count} leave applications that need approval <br />";
                body += $"Please go to <a href=\"{url}/leaveApprove\">this link</a> for approving.";

                var needSend = Convert.ToBoolean(configuration.GetSection("Quartz:UnapprovedNotify:NeedNotify").Value);
                if (needSend)
                {
                    mailService.Send(fromEmail, toEmail, title, body);
                }
            });
        }

        public class NotifyUser
        {
            public Sys_Users Users { get; set; }
            public int count { get; set; }
        }

    }
}
