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
    public class Sys_Menus:BaseEntity
    {
        public string RouteName { get; set; }
        public string MenuName { get; set; }
        public string? ParentID { get; set; }
        public string RoutePath { get; set; }
        public string Component { get; set; }
        public int MenuType { get; set; }
        public int Visible { get; set; }
        public string? Permissions { get; set; }
        public string? Icon { get; set; }
        public int Sort { get; set; }
        public string? Remark { get; set; }
        public bool KeepAlive { get; set; }


        [NotMapped]
        public List<Sys_Menus> Children
        {
            get { return this.children; }
            set { this.children = value; }
        }

        private List<Sys_Menus> children = new List<Sys_Menus>();

        public void SetChildren(List<Sys_Menus> menus)
        {
            children.AddRange(menus);
        }
    }
}
