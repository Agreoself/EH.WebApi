using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Atd_OtherRelated : BaseEntity
    {
        public string RequestID { get;set; }
        public string UserId { get; set; }
        public DateTime BornDate { get; set; }
        public virtual string Attachment { get; set; }
        public string Description { get; set; }
        public int CurrentState { get; set; }
    }
}
