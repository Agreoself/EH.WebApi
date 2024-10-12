using EH.System.Commons;
using EH.System.Models.Dtos;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.Service.Interface.Sys
{
    public interface IAucRecordService : IBaseService<Auc_Record>
    {
        public byte[] Export(string productId);

        public Task<bool> PlaceBidAsync(Auc_Record entity);
        public Decimal GetCurrentPrice(string productId);
    }
}
