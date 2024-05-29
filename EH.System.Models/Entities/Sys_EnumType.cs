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
    public class Sys_EnumType : BaseEntity
    {
        public string EnumName { get; set; }
        public string? EnumCode { get; set; } 
        public string? Description { get; set; } 
        
    }
}
