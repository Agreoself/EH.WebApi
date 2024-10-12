using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Auc_ProductDto
    {
        public Guid ID { get; set; }
        public string SkuCode { get; set; }
        public string SkuName { get; set; }
        public decimal BasePrice { get; set; }
        public decimal SinglePrice { get; set; }
        public int DelayPeriod { get; set; }
        public string ActivityId { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public IFormFileCollection ImageFile { get; set; }
        public string? Description { get; set; }
    }
}
