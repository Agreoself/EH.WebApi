using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Atd_FormAndProcess
    {

        public string LeaveId { get; set; }
        public string UserId { get; set; }

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
        public bool IsTreated { get; set; }

        public string ProcessID { get; set; }
        public string FormID { get; set; }

    }
}
