using EH.System.Commons;
using Microsoft.Extensions.Configuration;
using System;
using System.Net;
using System.Net.Mail;


namespace EH.System.Commons
{
    public interface IMailService : IScoped
    {
        void Send(string formEmail, string toEmail, string subject, string body, List<string> cc = null);

    }
    public class EmailService : IMailService
    {
        private readonly string smtp;
        private readonly int port;
        private readonly LogHelper logHelper;

        public EmailService(IConfiguration configuration, LogHelper logHelper)
        {
            // 从配置文件中获取发件人邮箱和密码

            var smtp = configuration.GetSection("EmailSettings:Host").Value.ToString();
            var port = Convert.ToInt32(configuration.GetSection("EmailSettings:Port").Value);

            if (string.IsNullOrEmpty(smtp))
                throw new ArgumentException("Missing smtp or port in config");

            this.smtp = smtp;
            this.port = port;
            this.logHelper = logHelper;
        }

        public void Send(string formEmail, string toEmail, string subject, string body, List<string> cc = null)
        {
            using (var client = new SmtpClient())
            {
                client.Port = port;
                client.Host = smtp;
                client.UseDefaultCredentials = true;

                //client.EnableSsl = true;
                var message = new MailMessage(formEmail, toEmail)
                {
                    Subject = subject,
                    //Body = body,
                    IsBodyHtml = true
                };
                AlternateView view = AlternateView.CreateAlternateViewFromString(body,null,"text/html");
                message.AlternateViews.Add(view);
                if (cc != null)
                {
                    foreach (var item in cc)
                    {
                        message.CC.Add(item);
                    }
                }
                //client.Credentials = new NetworkCredential(_fromEmail, _fromPassword);
                try
                {
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    logHelper.LogError("SendEmail error:" + ex.ToString());
                }

            }
        }
    }
}