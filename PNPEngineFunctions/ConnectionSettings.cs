using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PNPEngineFunctions
{
    public class ConnectionSettings
    {
        public string Login { get; set; }
        public string Password { get; set; }
        public string SiteUrl { get; set; }
        public bool UseAppAuthentication { get; set; }
    }
}
