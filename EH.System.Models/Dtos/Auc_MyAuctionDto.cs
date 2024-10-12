using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Auc_MyAuctionDto
    {
        public string ID { get; set; }
        public string Code { get; set; }
        public string Name { get; set; } 
        //get
        public decimal Price { get; set; }
        public DateTime Time { get; set; }
        //join
        public decimal Bid { get; set; }
        public decimal CurrentPrice { get; set; }
        public string ProductLifecycle { get; set; } 
        public string BidLifecycle { get; set; }
 
    }
}
