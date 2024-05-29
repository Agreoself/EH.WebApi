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
    public class Atd_LeaveForm:BaseEntity
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


    }
}
