using EH.Service.Interface.Attendance;
using EH.System.Commons;
using EH.System.Models.Dtos;
using Microsoft.AspNetCore.Cors.Infrastructure;
using Microsoft.Extensions.DependencyInjection;

namespace EH.System.HostServices
{
    public class EmailListenerHostedService : BackgroundService
    {
        private readonly IEmailListenerService emailListener;
        private readonly IConfiguration configuration;
        private readonly IServiceScopeFactory serviceScopeFactory;

        private readonly LogHelper log;

        public EmailListenerHostedService(IEmailListenerService emailListener, LogHelper log, IConfiguration configuration, IServiceScopeFactory serviceScopeFactory)
        {
            this.emailListener = emailListener;
            this.log = log;
            this.configuration = configuration;
            this.serviceScopeFactory = serviceScopeFactory;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    var waitAuditList = await emailListener.StartListening(stoppingToken);
                    foreach (var auditReq in waitAuditList)
                    {
                        using var scope = serviceScopeFactory.CreateScope();
                        var formService = scope.ServiceProvider.GetRequiredService<IAtdLeaveFormService>();
                        var auditRes = formService.ApproveByEmail(auditReq);
                        if (auditRes)
                        {
                            MoveInfo moveInfo = new()
                            {
                                MessageID = auditReq.MessageId,
                                Email = auditReq.MoveEmail,
                            };
                            var res = emailListener.MoveToArchive(moveInfo);
                        }

                    }
                }
                catch (Exception ex)
                {
                    // 处理异常，可以记录日志等
                    log.LogError($"An error occurred while listening for emails: {ex.Message}");
                }
                var seconds = Convert.ToInt32(configuration.GetSection("EmailIMAPSetting:Seconds").Value);
                // 等待一段时间后重新尝试
                await Task.Delay(TimeSpan.FromSeconds(seconds), stoppingToken);
            }
        }
    }
}
