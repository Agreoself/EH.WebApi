using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Common
{
    //[NotMapped]
    public abstract class BaseEntity
    {
        public Guid ID { get; set; }
        public DateTime CreateDate { get; set; }
        public string CreateBy { get; set; }
        public DateTime? ModifyDate { get; set; }
        public string ModifyBy { get; set; }
        public int Status { get; set; }
        public bool IsDeleted { get; set; }
    }
}
