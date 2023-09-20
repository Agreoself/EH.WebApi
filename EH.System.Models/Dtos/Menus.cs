using EH.System.Models.Common;
using EH.System.Models.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EH.System.Models.Dtos
{
    public class Menus : BaseEntity
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
        public string? Remark { get; set; }

        public int Sort { get; set; }

        public List<Menus> Children
        {
            get { return this.children; }
            set { this.children = value; }
        }

        private List<Menus> children = new List<Menus>();

        public void SetChildren(List<Menus> menus)
        {
            children.AddRange(menus);
        } 
    }
}
