using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using DateTimeConverter = EH.System.Models.Common.DateTimeConverter;

namespace EH.System.Models.Entities
{
    public class Atd_LeaveForm : BaseEntity
    {
        public string LeaveId { get; set; }

        public string UserId { get; set; }
        public string? Department { get; set; }

        public string LeaveType { get; set; }
        [JsonConverter(typeof(DateTimeConverter))]
        public DateTime StartDate { get; set; }
        [JsonConverter(typeof(DateTimeConverter))]
        public DateTime EndDate { get; set; }
        public decimal? TotalHours { get; set; }
        public string? Reason { get; set; }
        public string? Attachment { get; set; }
        public int CurrentState { get; set; }
        public bool IsCancel { get; set; }
        public int? CancelState { get; set; }
        public bool IsTreated { get; set; }


        public async Task<string> SetAttachment()
        {
            List<string> fileNames = new();

            var path = Directory.GetCurrentDirectory();
            var root = path + "\\HRVacation";
            var folder = root + "\\" + LeaveId + "\\";
            try
            {
                if (!string.IsNullOrEmpty(Attachment))
                {
                    var listBase64 = Attachment.Split("|pic|");
                    foreach (var bs64 in listBase64)
                    {
                        var filePath = folder;

                        if (bs64.StartsWith("data:"))
                        {
                            var fileBase = bs64.Split(";base64,");
                            var base64 = fileBase[1];
                            var ext = fileBase[0].Split("/")[1];
                            var fileBytes = Convert.FromBase64String(base64);

                            filePath = Path.Combine(folder, DateTime.Now.ToString("yyyyMMddHHmmssfff") + "." + ext);

                            if (!Directory.Exists(folder))
                            {
                                Directory.CreateDirectory(folder);
                            }

                            // Using File.WriteAllBytesAsync to write the file
                            await File.WriteAllBytesAsync(filePath, fileBytes);
                            fileNames.Add(filePath);
                        }
                        else
                        {
                            var originUrl = folder+ bs64.Split(LeaveId)[1];
                            fileNames.Add(originUrl);
                        }
                    }
                    Attachment = string.Join(",", fileNames);
                }
                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }

    }
}
