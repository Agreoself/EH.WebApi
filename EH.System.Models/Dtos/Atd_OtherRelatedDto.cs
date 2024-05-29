using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    internal class Atd_OtherRelatedDto
    {
        public Guid ID { get; set; }
        public string RequestID { get; set; }
        public string UserId { get; set; }
        public DateTime BornDate { get; set; }
        public string Description { get; set; }
        public int CurrentState { get; set; }

        public DateTime CreateDate { get; set; }
        public string CreateBy { get; set; }
        public DateTime? ModifyDate { get; set; }
        public string ModifyBy { get; set; }
        public int Status { get; set; }
        public bool IsDeleted { get; set; }
    }
}
