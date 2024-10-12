using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Auc_Activity : BaseEntity
    {
        public string AuctionCode { get; set; }
        public string AuctionName { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string? Description { get; set;}

        public int Lifecycle { get; set; }

        public void SetLifecycle()
        {
            Lifecycle = DateTime.Now < StartTime ? 0 : EndTime < DateTime.Now ? -1 : 1;
        }
    }
}
