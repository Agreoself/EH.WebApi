using EH.System.Models.Common;
using EH.System.Models.Dtos;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Entities
{
    public class Sys_EnumItem : BaseEntity
    {
        public string EnumTypeId { get; set; }
        public string? Text { get; set; }
        public string? Value { get; set; }
        public string? ParentId { get; set; }
        public string? Description { get; set; }
        public int Sort { get; set; } 
        
    }
}
