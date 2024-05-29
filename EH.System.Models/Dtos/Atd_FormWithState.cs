using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using DateTimeConverter = EH.System.Models.Common.DateTimeConverter;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Atd_FormWithState
    {
        public string? CurrentStep { get; set; }
        public string? CurrentOwner { get; set; }
        public string LeaveId { get; set; }

        public string UserId { get; set; }
        public string FullName { get; set; }
        public string? ChineseName { get; set; }
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

        public Guid ID { get; set; }
        public DateTime CreateDate { get; set; }
        public string CreateBy { get; set; }
        public DateTime? ModifyDate { get; set; }
        public string ModifyBy { get; set; }
        public int Status { get; set; }
        public bool IsDeleted { get; set; }
    }
}
