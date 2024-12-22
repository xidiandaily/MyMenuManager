using System.Collections.Generic;

namespace MyMenuManager.Models
{
    public class MenuConfig
    {
        public string Title { get; set; }
        public string Target { get; set; }
        public string Cmd { get; set; }
        public List<MenuConfig> Submenu { get; set; }
    }
} 