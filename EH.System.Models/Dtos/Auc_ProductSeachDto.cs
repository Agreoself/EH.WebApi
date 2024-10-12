using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Auc_ProductSeachDto
    {
        public PageRequest<Auc_Product> pageRequest { get; set; }
        public string date {  get; set; }

        public string nameOrSummary {  get; set; }
    }
}
