using EH.System.Commons;
using EH.System.Models.Entities;
using EH.Repository.Interface.Attendance;
using EH.Service.Interface.Attendance;
using EH.Repository.Interface.Sys;
using System.Net.Sockets;
using System.Security.Cryptography.X509Certificates;
using NPOI.SS.Formula.Functions;
using Microsoft.Extensions.Configuration;
using NPOI.XWPF.UserModel;
using EH.System.Models.Dtos;
using System.Security.Policy;
using EH.Service.Interface.Sys;
using System;
using EH.Repository.Implement;

namespace EH.Service.Implement.Attendance
{
    public class AtdOtherRelatedService : BaseService<Atd_OtherRelated>, IAtdOtherRelatedService, ITransient
    {
        private readonly LogHelper logHelper;
        private readonly IAtdOtherRelatedRepository repository;
        private readonly IAtdLeaveBalanceService balanceService;

        private readonly ISysUsersRepository userRepository;
        private readonly IMailService mailService;
        private readonly IConfiguration configuration;
        public AtdOtherRelatedService(IAtdOtherRelatedRepository repository, IAtdLeaveBalanceService balanceService, IMailService mailService, IConfiguration configuration, ISysUsersRepository userRepository, LogHelper logHelper) : base(repository, logHelper)
        {
            this.logHelper = logHelper;
            this.repository = repository;
            this.balanceService = balanceService;
            this.mailService = mailService;
            this.configuration = configuration;
            this.userRepository = userRepository;
        }

        public override List<Atd_OtherRelated> GetPageList(PageRequest<Atd_OtherRelated> request, out int totalCount)
        {
            var whereCondision = request.GetWhere().Compile();
            var orderCondision = request.GetOrder().Compile();

            var queryList= repository.Where(whereCondision, orderCondision, request.PageIndex, request.PageSize, out totalCount, isDesc: request.isDesc);
            var list = queryList.Select(e => new Atd_OtherRelated
            {
                ID= e.ID,
                RequestID= e.RequestID,
                UserId=e.UserId,
                BornDate=e.BornDate,
                CurrentState=e.CurrentState,
                Description=e.Description,
                CreateBy=e.CreateBy,
                CreateDate=e.CreateDate,
                IsDeleted =e.IsDeleted,
                Status=e.Status,
                ModifyBy  =e.ModifyBy,
                ModifyDate=e.ModifyDate,
                Attachment=e.Attachment==null?"":"yes"
            }).ToList();
            return list;
        }

        public override Atd_OtherRelated Insert(Atd_OtherRelated entity, bool isSave = true)
        {
            var needSend = Convert.ToBoolean(configuration.GetSection("EmailNotify:IsRequire").Value);
            var user = userRepository.FirstOrDefault(i => i.UserName == entity.UserId);

            var url = configuration.GetSection("EmailNotify:Url").Value;
            string toEmail = configuration.GetSection("EmailNotify:HrEmail").Value;
            var Hr = toEmail.Split(".")[0];

            string title = $"New parental leave application from {user.FullName}";
            string body = $"Dear {Hr} ,<br />";
            body += $"A parental leave application {entity.RequestID} requires your approval. <br />";
            body += $"Please go to {url}/otherRelated?requestId={entity.RequestID} to view details.";
            var fromEmail = user.Email;

            if (needSend)
            {
                mailService.Send(fromEmail, toEmail, title, body);
            }

            return base.Insert(entity, isSave);
        }

        public override bool Update(Atd_OtherRelated entity)
        {
            if (entity.CurrentState == 2)
            {
                var res = balanceService.CalculateParentalAndBreastfeeding(entity.UserId, entity.BornDate);

                if (res)
                {
                    var needSend = Convert.ToBoolean(configuration.GetSection("EmailNotify:IsRequire").Value);
                    var user = userRepository.FirstOrDefault(i => i.UserName == entity.UserId);

                    var url = configuration.GetSection("EmailNotify:Url").Value;
                    string hrEmail = configuration.GetSection("EmailNotify:HrEmail").Value;
                    var Hr = hrEmail.Split(".")[0];

                    var title = "Your leave request has been approved";
                    var body = $"Dear {user.FullName} <br/>";
                    body += $"Your parental leave application {entity.RequestID} has been approved by {Hr} <br/>";
                    body += $"Please go to {url}/otherRelated?requestId={entity.RequestID} for checking.";

                    var toEmail = user.Email;

                    if (needSend)
                    {
                        mailService.Send(hrEmail, toEmail, title, body);
                    }

                }
                return base.Update(entity);
            }
            else if (entity.CurrentState == 3)
            {
                var needSend = Convert.ToBoolean(configuration.GetSection("EmailNotify:IsRequire").Value);
                var user = userRepository.FirstOrDefault(i => i.UserName == entity.UserId);

                var url = configuration.GetSection("EmailNotify:Url").Value;
                string hrEmail = configuration.GetSection("EmailNotify:HrEmail").Value;
                var Hr = hrEmail.Split(".")[0];

                var title = "Your leave request has been rejected"; 
                var body = $"Dear {user.FullName} <br/>";
                body += $"Sorry, Your parental leave application {entity.RequestID} has been rejected by {Hr}, please reapply.";

                var toEmail = user.Email;

                if (needSend)
                {
                    mailService.Send(hrEmail, toEmail, title, body);
                }
                return base.Update(entity);
            }
            else
            { 
                return base.Update(entity);

            }
        }

    }
}
