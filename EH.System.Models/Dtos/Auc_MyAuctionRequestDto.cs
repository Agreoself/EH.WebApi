using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Auc_MyAuctionRequestDto
    {
        public int pageIndex { get; set; }
        public int pageSize { get; set; }
        public string userId { get; set; }
        public bool isEnd { get; set; }
    }
}
