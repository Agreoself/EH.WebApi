using EH.System.Models.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Auc_Product : BaseEntity
    {
        public string ActivityId { get; set; }
        public string SkuCode { get; set; }
        public string SkuName { get; set; }
        public string Category { get; set; }
        public string? Defects { get; set; }
        public string? Summary { get; set; }
        public decimal BasePrice { get; set; }
        public decimal SinglePrice { get; set; }
        public int DelayPeriod { get; set; }

        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string? Images { get; set; }
        public string? Description { get; set; }
        public int Lifecycle { get; set; }
        public decimal CurrentPrice { get; set; }

        public void SetLifecycle()
        {
            Lifecycle = DateTime.Now < StartTime ? 0 :EndTime<DateTime.Now?-1: 1;
        }
    }
}
