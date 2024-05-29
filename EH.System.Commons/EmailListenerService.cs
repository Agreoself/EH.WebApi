using EH.System.Models.Dtos;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using NPOI.SS.Formula.Functions;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Text;

namespace EH.System.Commons
{

    public interface IEmailListenerService : ISingleton
    {
        //Task StartListening(string email, CancellationToken cancellationToken = default);
        Task<List<Atd_ApproveByEmail>> StartListening(CancellationToken cancellationToken = default);
        Task<bool> MoveToArchive(MoveInfo info);

        //Task StartListening(string email);
        //Task StopListening();

    }
    public class EmailListenerService : IEmailListenerService
    {
        private readonly LogHelper logHelper;
        private string accessToken;
        private readonly IConfiguration configuration;

        public EmailListenerService(LogHelper logHelper, IConfiguration configuration)
        {
            this.logHelper = logHelper;
            this.configuration = configuration;
        }

        public async Task<List<Atd_ApproveByEmail>> StartListening(CancellationToken cancellationToken = default)
        {
            try
            {
                List<Atd_ApproveByEmail> waitAuditList = new List<Atd_ApproveByEmail>();
                var email = configuration.GetSection("EmailIMAPSetting:Email").Value.ToString();
                accessToken = await GetAccessToken();
                if (string.IsNullOrEmpty(accessToken))
                {
                    return waitAuditList;
                }
                var inboxMail = await GetInboxMail();
                if (inboxMail != null)
                {
                    if (!string.IsNullOrEmpty(Convert.ToString(inboxMail["value"])))
                    {
                        foreach (JToken result in inboxMail["value"])
                        {
                            if (!result["categories"].Any())
                            {
                             
                                var subject = result["subject"].ToString();
                                var leaveId = subject.Split('-')[0].Split("=")[1];
                                var auditResult = subject.Split('-')[1].Split("=")[1];

                                var body = result["bodyPreview"].ToString();
                                body = body.Contains("Comment:") ? body.Replace("Comment:", "") : body;
                                body = body.Contains("Comment：") ? body.Replace("Comment：", "") : body;

                                Atd_ApproveByEmail approveByEmail = new()
                                {
                                    MoveEmail=email,
                                    FromEmail = result["from"]["emailAddress"]["address"].ToString(),
                                    LeaveId = leaveId,
                                    Comment = body,
                                    Result = auditResult,
                                    MessageId = result["id"].ToString(),
                                };

                                waitAuditList.Add(approveByEmail);
                            }
                        }
                    }
                }
                return waitAuditList;
            }
            catch (Exception ex)
            {
                logHelper.LogError("Listening HRVac@ehealth.com inbox Exception: " + ex.Message);
                return new List<Atd_ApproveByEmail>();
            }


        }

        async Task<string> GetAccessToken()
        {
            // 这里假设使用 Microsoft 的身份验证库获取访问令牌
            var clientSecret = configuration.GetSection("EmailIMAPSetting:Secret").Value.ToString();
            var clientId = configuration.GetSection("EmailIMAPSetting:Appid").Value.ToString();
            var tenantId = configuration.GetSection("EmailIMAPSetting:TenantId").Value.ToString();
            try
            {
                using HttpClient httpClient = new();
                var request = new HttpRequestMessage(HttpMethod.Post,
        $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token");
                var httpBody =
                    new StringContent(
                        $"grant_type=client_credentials&client_id={clientId}&scope=https%3A%2F%2Fgraph.microsoft.com%2F.default&client_secret={clientSecret}");
                httpBody.Headers.ContentType.MediaType = "application/x-www-form-urlencoded";
                request.Content = httpBody;

                var jo = new JObject();
                var tokenTask = await httpClient.SendAsync(request);
                // 异步操作
                var tokenRep = await tokenTask.Content.ReadAsStringAsync();
                jo = JObject.Parse(tokenRep);
                if (!string.IsNullOrEmpty(Convert.ToString(jo["access_token"])))
                {
                    return Convert.ToString(jo["access_token"]);
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                logHelper.LogError("Microsoft Get AccessToken Exception: " + ex.Message);
                return "";
            }
        }

        public async Task<JObject> GetInboxMail()
        {
            var email = configuration.GetSection("EmailIMAPSetting:Email").Value.ToString();
            using HttpClient httpClient = new();
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            HttpResponseMessage response = await httpClient.GetAsync($"https://graph.microsoft.com/v1.0/users/{email}/mailFolders/inbox/messages");

            if (response.IsSuccessStatusCode)
            {
                string responseBody = await response.Content.ReadAsStringAsync();
                JObject json = JObject.Parse(responseBody); //在这里会等待task返回。
                return json;
            }
            else
            {
                logHelper.LogError("Get HRVac@ehealth.com inbox fail: " + response.ReasonPhrase);
                return null;
            }
        }

        /// <summary>
        /// 将邮件移动到archive文件夹
        /// </summary>
        /// <param name="info"></param>
        /// <returns></returns>
        public async Task<bool> MoveToArchive(MoveInfo info)
        {
            // 使用访问令牌发送请求获取邮件
            using HttpClient httpClient = new();
            httpClient.DefaultRequestHeaders.Add("Authorization", "Bearer " + accessToken);
            JObject jObject = new()
            {
                ["destinationId"] = "archive"
            };
            var httpBody = new StringContent(jObject.ToString(), Encoding.UTF8, "application/json");
            HttpResponseMessage response = await httpClient.PostAsync($"https://graph.microsoft.com/v1.0/users/{info.Email}/messages/{info.MessageID}/move", httpBody);

            if (response.IsSuccessStatusCode)
            {
                return true;
            }
            else
            {
                logHelper.LogError("Move HRVac@ehealth.com inbox to Archive fail: " + response.ReasonPhrase);
                return false;
            }
        }

        //public async Task StartListening(string email)
        //{
        //    try
        //    {
        //        var imapServer = configuration.GetSection("EmailIMAPSetting:Host").Value.ToString();
        //        var imapPort = Convert.ToInt32(configuration.GetSection("EmailIMAPSetting:Port").Value);
        //        // 连接到 B 邮箱的 IMAP 服务器
        //        await _client.ConnectAsync(imapServer, imapPort); // 替换为 B 邮箱的 IMAP 服务器地址和端口
        //        await _client.AuthenticateAsync(email, null);

        //        // 打开收件箱
        //        var inbox = await _client.Inbox.OpenAsync(FolderAccess.ReadOnly);

        //        // 启动监听新邮件
        //        _client.Inbox.CountChanged += async (sender, e) =>
        //        {
        //            // 获取新邮件的唯一标识符
        //            var uids = _client.Inbox.Search(SearchQuery.New);

        //            // 遍历新邮件
        //            foreach (var uid in uids)
        //            {
        //                var message = await _client.Inbox.GetMessageAsync(uid);

        //                // 在此处执行你想要做的操作，例如打印邮件内容
        //                logHelper.LogInfo($"New Email Received: {message.Subject}");
        //                logHelper.LogInfo($"From: {message.From}");
        //                logHelper.LogInfo($"Body: {message.TextBody}");

        //                // 这里可以执行其他操作，比如解析邮件内容，触发其他事件等
        //            }
        //        };
        //    }
        //    catch (Exception ex)
        //    {
        //        logHelper.LogError("出现异常: " + ex.Message);
        //    }
        //}


    }
}
