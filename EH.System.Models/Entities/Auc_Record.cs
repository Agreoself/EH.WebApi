using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Auc_Record : BaseEntity
    {
        public DateTime BidTime { get; set; }
        public string UserId { get; set; }
        public int Lifecycle { get; set; }
        public string ProductId { get; set; }
        public decimal BidPrice { get; set; } 
  
    }
}
